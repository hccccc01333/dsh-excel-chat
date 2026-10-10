import { readFile } from 'node:fs/promises'
import ExcelJS from 'exceljs'
import { t, listJoin, listSeparator } from './i18n.ts'
import { normalizeCellId } from './score.ts'
import { readWorkbookCells, stripPivotTableParts } from './workbook.ts'

/** A deterministic assertion over one workbook cell and its presentation. */
export interface CellAssertion {
  /** Cell id, such as `订单!B3`. */
  id: string
  /** Exact serialized cell content; `null` requires an empty or absent cell. */
  expect?: string | null
  /** Required serialized cell-content prefix, useful for formulas. */
  startsWith?: string
  /** Required foreground fill color, with or without an ARGB alpha prefix. */
  fill?: string
  /** Required bold state. */
  bold?: boolean
  /** Required Excel number format. */
  numberFormat?: string
  /** Required wrap-text state. */
  wrapText?: boolean
  /** Required horizontal alignment. */
  hAlign?: string
}

/**
 * A deterministic assertion over a sheet's view and layout.
 *
 * These operations change nothing a cell assertion can see. Freezing panes, hiding a
 * row, setting a column width — the cell contents are untouched, so the corpus could not
 * assert them at all and 38 of the 77 operations had no end-to-end coverage. This is the
 * missing subject: the sheet itself.
 *
 * Discriminated by `sheet` rather than a `kind` field, because the two variants are
 * already distinguishable and an extra tag would have to be threaded through every
 * existing assertion.
 */
export interface SheetAssertion {
  /** Sheet name. */
  sheet: string
  /** Required frozen-pane cell (`A2` freezes the first row); `null` requires no freeze. */
  frozen?: string | null
  /** Rows that must be hidden. */
  hiddenRows?: number[]
  /** Columns that must be hidden. */
  hiddenColumns?: string[]
  /** Required column widths, keyed by column letter. */
  columnWidth?: Record<string, number>
  /** Required row heights, keyed by row number. */
  rowHeight?: Record<string, number>
}

export type WorkbookAssertion = CellAssertion | SheetAssertion

/** The result for one deterministic workbook assertion. */
export interface WorkbookAssertionResult {
  id: string
  passed: boolean
  detail: string
}

/** Evidence returned by the deterministic verifier. */
export interface WorkbookVerification {
  achieved: boolean
  passed: number
  total: number
  failures: string[]
  assertions: WorkbookAssertionResult[]
  reason: string
}

/**
 * Evaluate workbook assertions without an LLM.
 *
 * Cell-content assertions use the normalized workbook cell map. Presentation
 * assertions load the workbook once and require every declared presentation
 * property on a check to match.
 */
export async function verifyWorkbookAssertions(
  path: string,
  assertions: WorkbookAssertion[],
): Promise<WorkbookVerification> {
  const cells = await readWorkbookCells(await readFile(path))
  // A sheet assertion always needs the parsed workbook; a cell assertion only does when
  // it asks about presentation. Narrow first, so `hasStyleAssertion` sees a cell.
  const needsWorkbook = assertions.some(
    (assertion) => isSheetAssertion(assertion) || hasStyleAssertion(assertion),
  )
  const workbook = needsWorkbook ? await loadWorkbook(path) : null
  const results = assertions.map((assertion) =>
    isSheetAssertion(assertion)
      ? evaluateSheetAssertion(assertion, workbook!)
      : evaluateAssertion(assertion, cells, workbook ? styleCellsOf(workbook) : null),
  )
  const failures = results.filter((result) => !result.passed).map((result) => result.detail)
  const passed = results.length - failures.length
  const achieved = results.length > 0 && failures.length === 0
  const reason = achieved
    ? t('确定性断言全部通过（{passed}/{total}）', { passed, total: results.length })
    : results.length === 0
      ? t('没有可执行的确定性断言')
      : t('确定性断言未全部通过（{passed}/{total}）：{failures}', {
          passed,
          total: results.length,
          failures: listJoin(failures.slice(0, 4), 'semicolon'),
        })
  return { achieved, passed, total: results.length, failures, assertions: results, reason }
}

/** A sheet assertion names a sheet; a cell assertion names a cell. */
function isSheetAssertion(assertion: WorkbookAssertion): assertion is SheetAssertion {
  return typeof (assertion as SheetAssertion).sheet === 'string' && !(assertion as CellAssertion).id
}

/**
 * Evaluate a sheet's view and layout.
 *
 * Every condition named is required, same as the cell path: an assertion that asks for a
 * frozen pane *and* hidden rows has to find both, or it reports what it actually saw.
 */
function evaluateSheetAssertion(
  assertion: SheetAssertion,
  workbook: ExcelJS.Workbook,
): WorkbookAssertionResult {
  const sheet = workbook.getWorksheet(assertion.sheet)
  if (!sheet) {
    return {
      id: assertion.sheet,
      passed: false,
      detail: t('{sheet} 不存在', { sheet: assertion.sheet }),
    }
  }

  const checks: Array<{ label: string; passed: boolean }> = []

  if (assertion.frozen !== undefined) {
    // ExcelJS keeps panes in the sheet's views; a freeze is `state: 'frozen'` plus the
    // split position, and an unfrozen sheet simply has no such view.
    const view = (sheet.views ?? []).find((entry) => entry.state === 'frozen') as
      | { state?: string; xSplit?: number; ySplit?: number; topLeftCell?: string }
      | undefined
    const actual = view?.topLeftCell ?? null
    checks.push({
      label: t('冻结窗格 期望 {expected} 实际 {actual}', {
        expected: assertion.frozen ?? t('无'),
        actual: actual ?? t('无'),
      }),
      passed: (actual ?? null) === (assertion.frozen ?? null),
    })
  }

  if (assertion.hiddenRows !== undefined) {
    for (const row of assertion.hiddenRows) {
      const actual = sheet.getRow(row).hidden === true
      checks.push({ label: t('第 {row} 行隐藏 期望是', { row }), passed: actual })
    }
  }

  if (assertion.hiddenColumns !== undefined) {
    for (const column of assertion.hiddenColumns) {
      const actual = sheet.getColumn(column).hidden === true
      checks.push({ label: t('{column} 列隐藏 期望是', { column }), passed: actual })
    }
  }

  if (assertion.columnWidth !== undefined) {
    for (const [column, width] of Object.entries(assertion.columnWidth)) {
      const actual = sheet.getColumn(column).width
      // ExcelJS round-trips widths with float noise; compare to a tenth of a unit.
      const passed = typeof actual === 'number' && Math.abs(actual - width) < 0.1
      checks.push({
        label: t('{column} 列宽 期望 {expected} 实际 {actual}', {
          column,
          expected: width,
          actual: actual ?? t('未设置'),
        }),
        passed,
      })
    }
  }

  if (assertion.rowHeight !== undefined) {
    for (const [row, height] of Object.entries(assertion.rowHeight)) {
      const actual = sheet.getRow(Number(row)).height
      const passed = typeof actual === 'number' && Math.abs(actual - height) < 0.1
      checks.push({
        label: t('第 {row} 行高 期望 {expected} 实际 {actual}', {
          row,
          expected: height,
          actual: actual ?? t('未设置'),
        }),
        passed,
      })
    }
  }

  if (checks.length === 0) {
    return { id: assertion.sheet, passed: false, detail: t('{sheet} 的表级断言没有声明任何条件', { sheet: assertion.sheet }) }
  }

  const failed = checks.filter((check) => !check.passed)
  return {
    id: assertion.sheet,
    passed: failed.length === 0,
    detail:
      failed.length === 0
        ? t('{sheet} 表级断言通过', { sheet: assertion.sheet })
        : t('{sheet} 表级断言未通过：{failures}', {
            sheet: assertion.sheet,
            failures: listJoin(failed.map((check) => check.label), 'semicolon'),
          }),
  }
}

/**
 * Evaluate every condition an assertion declares and require all of them.
 *
 * The earlier version returned as soon as one kind of condition matched, so
 * `{ expect, bold }` verified the value and silently ignored the style, and
 * `{ startsWith, fill }` ignored the fill. An assertion that names several
 * conditions is a conjunction — that is the whole point of naming them — and
 * checking only the first one produced exactly the false "achieved" this verifier
 * exists to prevent.
 */
function evaluateAssertion(
  assertion: CellAssertion,
  cells: Record<string, string>,
  styleCells: Map<string, ExcelJS.Cell> | null,
): WorkbookAssertionResult {
  const normalized = normalizeCellId(assertion.id)
  const actual = cells[normalized] ?? cells[findKey(cells, normalized) ?? '']
  const cell = styleCells?.get(normalized)

  const checks: Array<{ label: string; passed: boolean }> = []
  if (assertion.expect !== undefined) {
    checks.push({
      label: t('值 期望 {expected} 实际 {actual}', {
        expected: formatValue(assertion.expect),
        actual: formatValue(actual),
      }),
      passed: assertion.expect === null
        ? actual === undefined || actual === ''
        : actual === assertion.expect,
    })
  }
  if (assertion.startsWith !== undefined) {
    checks.push({
      label: t('前缀 期望 {expected} 实际 {actual}', {
        expected: formatValue(assertion.startsWith),
        actual: formatValue(actual),
      }),
      passed: typeof actual === 'string' && actual.startsWith(assertion.startsWith),
    })
  }
  if (assertion.fill !== undefined) {
    checks.push({ label: t('填充色 期望 {value}', { value: assertion.fill }), passed: colorMatches(cell, assertion.fill) })
  }
  if (assertion.bold !== undefined) {
    checks.push({ label: t('加粗 期望 {value}', { value: assertion.bold }), passed: (cell?.font?.bold ?? false) === assertion.bold })
  }
  if (assertion.numberFormat !== undefined) {
    checks.push({ label: t('数字格式 期望 {value}', { value: assertion.numberFormat }), passed: cell?.numFmt === assertion.numberFormat })
  }
  if (assertion.wrapText !== undefined) {
    checks.push({ label: t('自动换行 期望 {value}', { value: assertion.wrapText }), passed: (cell?.alignment?.wrapText ?? false) === assertion.wrapText })
  }
  if (assertion.hAlign !== undefined) {
    checks.push({ label: t('水平对齐 期望 {value}', { value: assertion.hAlign }), passed: cell?.alignment?.horizontal === assertion.hAlign })
  }

  const failed = checks.filter((check) => !check.passed)
  const passed = checks.length > 0 && failed.length === 0
  return {
    id: assertion.id,
    passed,
    detail: passed
      ? t('{id} 已满足全部 {count} 项要求', { id: assertion.id, count: checks.length })
      : t('{id} 未满足：{checks}', { id: assertion.id, checks: listJoin(failed.map((check) => check.label), 'semicolon') }),
  }
}

function hasStyleAssertion(assertion: CellAssertion): boolean {
  return assertion.fill !== undefined
    || assertion.bold !== undefined
    || assertion.numberFormat !== undefined
    || assertion.wrapText !== undefined
    || assertion.hAlign !== undefined
}

function colorMatches(cell: ExcelJS.Cell | undefined, expected: string): boolean {
  const actual = cell?.fill?.type === 'pattern'
    ? (cell.fill.fgColor as { argb?: string } | undefined)?.argb
    : undefined
  return actual !== undefined && actual.toUpperCase().endsWith(expected.toUpperCase())
}

function formatValue(value: string | null | undefined): string {
  return value === undefined ? t('缺失') : JSON.stringify(value)
}

async function loadWorkbook(path: string): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(stripPivotTableParts(await readFile(path)) as any)
  return workbook
}

/** Index every non-empty cell of an already-loaded workbook by its normalized id. */
function styleCellsOf(workbook: ExcelJS.Workbook): Map<string, ExcelJS.Cell> {
  const cells = new Map<string, ExcelJS.Cell>()
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        cells.set(normalizeCellId(`${sheet.name}!${cell.address}`), cell)
      })
    })
  })
  return cells
}

function findKey(cells: Record<string, string>, normalized: string): string | undefined {
  return Object.keys(cells).find((key) => normalizeCellId(key) === normalized)
}
