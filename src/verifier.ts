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
  /** Required protection state. */
  protected?: boolean
  /** Required tab colour, with or without an ARGB alpha prefix. */
  tabColor?: string
  /** Required autofilter range (`A1:B1`); `null` requires no filter. */
  autoFilter?: string | null
  /** Required cell comments, keyed by cell id, valued by the comment text. */
  comments?: Record<string, string>
  /** Required hyperlinks, keyed by cell id, valued by the URL. */
  hyperlinks?: Record<string, string>
  /** Required row outline levels, keyed by row number. */
  outlineRows?: Record<string, number>
  /** Required column outline levels, keyed by column letter. */
  outlineColumns?: Record<string, number>
  /** Required number of embedded images. */
  images?: number
  /** Table names that must exist. */
  tables?: string[]
  /** Required data-validation types, keyed by cell id. */
  dataValidation?: Record<string, string>
  /** Required page orientation, e.g. `landscape`. */
  pageOrientation?: string
  /** Required odd-page header string, as written (Excel codes and all). */
  oddHeader?: string
  /** Required gridline visibility. */
  showGridLines?: boolean
  /** Required zoom percentage. */
  zoom?: number
  /** Required sheet state. */
  state?: 'visible' | 'hidden' | 'veryHidden'
  /** Merged ranges that must exist, as written (`A1:B1`). */
  merged?: string[]
}

/**
 * A deterministic assertion over the workbook itself: sheet order and document
 * properties.
 *
 * These belong to no single sheet. `moveSheet` changes the order of the tab strip and
 * `setWorkbookProperties` writes the metadata Excel shows in File → Info; neither is
 * visible from any sheet, which is why neither had coverage.
 */
export interface WorkbookLevelAssertion {
  /** Discriminator. A sheet assertion names a sheet, a cell assertion a cell. */
  workbook: true
  /** Required sheet order, by name. */
  sheetOrder?: string[]
  /** Required document properties. */
  properties?: { title?: string; creator?: string }
}

export type WorkbookAssertion = CellAssertion | SheetAssertion | WorkbookLevelAssertion

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
    (assertion) => isWorkbookAssertion(assertion) || isSheetAssertion(assertion) || hasStyleAssertion(assertion),
  )
  const workbook = needsWorkbook ? await loadWorkbook(path) : null
  const results = assertions.map((assertion) =>
    isWorkbookAssertion(assertion)
      ? evaluateWorkbookAssertion(assertion, workbook!)
      : isSheetAssertion(assertion)
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

function isWorkbookAssertion(assertion: WorkbookAssertion): assertion is WorkbookLevelAssertion {
  return (assertion as WorkbookLevelAssertion).workbook === true
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

  if (assertion.protected !== undefined) {
    const protection = (sheet as unknown as { sheetProtection?: unknown }).sheetProtection
    const actual = protection !== undefined && protection !== null
    checks.push({
      label: t('保护 期望 {expected} 实际 {actual}', {
        expected: assertion.protected ? t('是') : t('否'),
        actual: actual ? t('是') : t('否'),
      }),
      passed: actual === assertion.protected,
    })
  }

  if (assertion.tabColor !== undefined) {
    const actual = (sheet.properties as { tabColor?: { argb?: string } } | undefined)?.tabColor?.argb ?? null
    checks.push({
      label: t('标签色 期望 {expected} 实际 {actual}', {
        expected: assertion.tabColor,
        actual: actual ?? t('无'),
      }),
      // Compare the tail so `FF0000` matches `FFFF0000`, the same rule the cell fill
      // uses: the alpha prefix is optional in both places.
      passed: actual !== null && actual.toUpperCase().endsWith(assertion.tabColor.toUpperCase()),
    })
  }

  if (assertion.autoFilter !== undefined) {
    const filter = sheet.autoFilter as string | { from?: string; to?: string } | undefined
    const actual = typeof filter === 'string' ? filter : filter ? `${filter.from}:${filter.to}` : null
    checks.push({
      label: t('筛选范围 期望 {expected} 实际 {actual}', {
        expected: assertion.autoFilter ?? t('无'),
        actual: actual ?? t('无'),
      }),
      passed: (actual ?? null) === (assertion.autoFilter ?? null),
    })
  }

  if (assertion.comments !== undefined) {
    for (const [cellId, text] of Object.entries(assertion.comments)) {
      const cell = cellIn(sheet, cellId)
      const note = cell?.note
      const actual = typeof note === 'string' ? note : note?.texts?.map((part) => part.text).join('') ?? null
      checks.push({
        label: t('{cell} 批注 期望 {expected} 实际 {actual}', {
          cell: cellId,
          expected: text,
          actual: actual ?? t('无'),
        }),
        passed: actual === text,
      })
    }
  }

  if (assertion.hyperlinks !== undefined) {
    for (const [cellId, url] of Object.entries(assertion.hyperlinks)) {
      const value = cellIn(sheet, cellId)?.value as { hyperlink?: string } | string | null
      const actual = typeof value === 'object' && value !== null ? value.hyperlink ?? null : null
      checks.push({
        label: t('{cell} 超链接 期望 {expected} 实际 {actual}', {
          cell: cellId,
          expected: url,
          actual: actual ?? t('无'),
        }),
        passed: actual === url,
      })
    }
  }

  if (assertion.outlineRows !== undefined) {
    for (const [row, level] of Object.entries(assertion.outlineRows)) {
      const actual = sheet.getRow(Number(row)).outlineLevel ?? 0
      checks.push({
        label: t('第 {row} 行分组层级 期望 {expected} 实际 {actual}', { row, expected: level, actual }),
        passed: actual === level,
      })
    }
  }

  if (assertion.outlineColumns !== undefined) {
    for (const [column, level] of Object.entries(assertion.outlineColumns)) {
      const actual = sheet.getColumn(column).outlineLevel ?? 0
      checks.push({
        label: t('{column} 列分组层级 期望 {expected} 实际 {actual}', { column, expected: level, actual }),
        passed: actual === level,
      })
    }
  }

  if (assertion.images !== undefined) {
    const actual = sheet.getImages?.().length ?? 0
    checks.push({
      label: t('图片数 期望 {expected} 实际 {actual}', { expected: assertion.images, actual }),
      passed: actual === assertion.images,
    })
  }

  if (assertion.tables !== undefined) {
    for (const name of assertion.tables) {
      const actual = sheet.getTable?.(name) !== undefined
      checks.push({ label: t('表格 {name} 存在 期望是', { name }), passed: actual })
    }
  }

  if (assertion.dataValidation !== undefined) {
    for (const [cellId, type] of Object.entries(assertion.dataValidation)) {
      const actual = cellIn(sheet, cellId)?.dataValidation?.type ?? null
      checks.push({
        label: t('{cell} 数据验证 期望 {expected} 实际 {actual}', {
          cell: cellId,
          expected: type,
          actual: actual ?? t('无'),
        }),
        passed: actual === type,
      })
    }
  }

  if (assertion.pageOrientation !== undefined) {
    const actual = sheet.pageSetup?.orientation ?? null
    checks.push({
      label: t('纸张方向 期望 {expected} 实际 {actual}', {
        expected: assertion.pageOrientation,
        actual: actual ?? t('无'),
      }),
      passed: actual === assertion.pageOrientation,
    })
  }

  if (assertion.oddHeader !== undefined) {
    const actual = sheet.headerFooter?.oddHeader ?? null
    checks.push({
      label: t('页眉 期望 {expected} 实际 {actual}', {
        expected: assertion.oddHeader,
        actual: actual ?? t('无'),
      }),
      passed: actual === assertion.oddHeader,
    })
  }

  if (assertion.showGridLines !== undefined) {
    const view = sheet.views?.[0] as { showGridLines?: boolean } | undefined
    const actual = view?.showGridLines ?? true
    checks.push({
      label: t('网格线 期望 {expected} 实际 {actual}', {
        expected: assertion.showGridLines ? t('显示') : t('隐藏'),
        actual: actual ? t('显示') : t('隐藏'),
      }),
      passed: actual === assertion.showGridLines,
    })
  }

  if (assertion.zoom !== undefined) {
    const view = sheet.views?.[0] as { zoomScale?: number } | undefined
    const actual = view?.zoomScale ?? 100
    checks.push({
      label: t('缩放 期望 {expected} 实际 {actual}', { expected: assertion.zoom, actual }),
      passed: actual === assertion.zoom,
    })
  }

  if (assertion.state !== undefined) {
    const actual = sheet.state ?? 'visible'
    checks.push({
      label: t('工作表状态 期望 {expected} 实际 {actual}', { expected: assertion.state, actual }),
      passed: actual === assertion.state,
    })
  }

  if (assertion.merged !== undefined) {
    // Exact match, not containment. `merged: []` has to mean "nothing is merged" —
    // written as a containment check first, an empty list checked nothing at all and
    // the assertion passed without looking, which is the failure mode this whole
    // assertion kind exists to prevent.
    const actual = new Set(
      ((sheet as unknown as { model?: { merges?: string[] } }).model?.merges ?? []).map((range) =>
        range.replace(/\$/g, '').toUpperCase(),
      ),
    )
    const wanted = new Set(assertion.merged.map((range) => range.replace(/\$/g, '').toUpperCase()))
    checks.push({
      label: t('合并区 期望 {expected} 实际 {actual}', {
        expected: [...wanted].sort().join(' ') || t('无'),
        actual: [...actual].sort().join(' ') || t('无'),
      }),
      passed: actual.size === wanted.size && [...wanted].every((range) => actual.has(range)),
    })
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

/** Evaluate sheet order and document properties. */
function evaluateWorkbookAssertion(
  assertion: WorkbookLevelAssertion,
  workbook: ExcelJS.Workbook,
): WorkbookAssertionResult {
  const checks: Array<{ label: string; passed: boolean }> = []

  if (assertion.sheetOrder !== undefined) {
    const actual = workbook.worksheets.map((sheet) => sheet.name)
    checks.push({
      label: t('工作表顺序 期望 {expected} 实际 {actual}', {
        expected: assertion.sheetOrder.join(' | '),
        actual: actual.join(' | '),
      }),
      passed:
        actual.length === assertion.sheetOrder.length &&
        actual.every((name, index) => name === assertion.sheetOrder![index]),
    })
  }

  if (assertion.properties?.title !== undefined) {
    const actual = workbook.title ?? null
    checks.push({
      label: t('文档标题 期望 {expected} 实际 {actual}', {
        expected: assertion.properties.title,
        actual: actual ?? t('无'),
      }),
      passed: actual === assertion.properties.title,
    })
  }

  if (assertion.properties?.creator !== undefined) {
    const actual = workbook.creator ?? null
    checks.push({
      label: t('文档作者 期望 {expected} 实际 {actual}', {
        expected: assertion.properties.creator,
        actual: actual ?? t('无'),
      }),
      passed: actual === assertion.properties.creator,
    })
  }

  if (checks.length === 0) {
    return { id: 'workbook', passed: false, detail: t('工作簿级断言没有声明任何条件') }
  }
  const failed = checks.filter((check) => !check.passed)
  return {
    id: 'workbook',
    passed: failed.length === 0,
    detail:
      failed.length === 0
        ? t('工作簿级断言通过')
        : t('工作簿级断言未通过：{failures}', {
            failures: listJoin(failed.map((check) => check.label), 'semicolon'),
          }),
  }
}

/** Resolve `A2` or `订单!A2` against an already-loaded sheet. */
function cellIn(sheet: ExcelJS.Worksheet, cellId: string): ExcelJS.Cell | undefined {
  const address = cellId.includes('!') ? cellId.split('!').pop()! : cellId
  try {
    return sheet.getCell(address)
  } catch {
    return undefined
  }
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
