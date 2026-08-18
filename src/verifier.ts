import { readFile } from 'node:fs/promises'
import ExcelJS from 'exceljs'
import { normalizeCellId } from './score.ts'
import { readWorkbookCells, stripPivotTableParts } from './workbook.ts'

/** A deterministic assertion over one workbook cell and its presentation. */
export interface WorkbookAssertion {
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
  const needsStyles = assertions.some((assertion) => hasStyleAssertion(assertion))
  const styleCells = needsStyles ? await loadStyleCells(path) : null
  const results = assertions.map((assertion) => evaluateAssertion(assertion, cells, styleCells))
  const failures = results.filter((result) => !result.passed).map((result) => result.detail)
  const passed = results.length - failures.length
  const achieved = results.length > 0 && failures.length === 0
  const reason = achieved
    ? `确定性断言全部通过（${passed}/${results.length}）`
    : results.length === 0
      ? '没有可执行的确定性断言'
      : `确定性断言未全部通过（${passed}/${results.length}）：${failures.slice(0, 4).join('；')}`
  return { achieved, passed, total: results.length, failures, assertions: results, reason }
}

function evaluateAssertion(
  assertion: WorkbookAssertion,
  cells: Record<string, string>,
  styleCells: Map<string, ExcelJS.Cell> | null,
): WorkbookAssertionResult {
  const normalized = normalizeCellId(assertion.id)
  const actual = cells[normalized] ?? cells[findKey(cells, normalized) ?? '']
  if (assertion.expect !== undefined) {
    const matches = assertion.expect === null
      ? actual === undefined || actual === ''
      : actual === assertion.expect
    return {
      id: assertion.id,
      passed: matches,
      detail: matches ? `${assertion.id} 已满足期望值` : `${assertion.id} 期望 ${formatValue(assertion.expect)}，实际 ${formatValue(actual)}`,
    }
  }
  if (assertion.startsWith !== undefined) {
    const matches = typeof actual === 'string' && actual.startsWith(assertion.startsWith)
    return {
      id: assertion.id,
      passed: matches,
      detail: matches ? `${assertion.id} 已满足前缀要求` : `${assertion.id} 期望以 ${formatValue(assertion.startsWith)} 开头，实际 ${formatValue(actual)}`,
    }
  }
  const cell = styleCells?.get(normalized)
  const checks: Array<{ label: string; passed: boolean }> = [
    ...(assertion.fill !== undefined ? [{ label: `填充色=${assertion.fill}`, passed: colorMatches(cell, assertion.fill) }] : []),
    ...(assertion.bold !== undefined ? [{ label: `加粗=${assertion.bold}`, passed: (cell?.font?.bold ?? false) === assertion.bold }] : []),
    ...(assertion.numberFormat !== undefined ? [{ label: `数字格式=${assertion.numberFormat}`, passed: cell?.numFmt === assertion.numberFormat }] : []),
    ...(assertion.wrapText !== undefined ? [{ label: `自动换行=${assertion.wrapText}`, passed: (cell?.alignment?.wrapText ?? false) === assertion.wrapText }] : []),
    ...(assertion.hAlign !== undefined ? [{ label: `水平对齐=${assertion.hAlign}`, passed: cell?.alignment?.horizontal === assertion.hAlign }] : []),
  ]
  const passed = checks.length > 0 && checks.every((check) => check.passed)
  const failedChecks = checks.filter((check) => !check.passed).map((check) => check.label)
  return {
    id: assertion.id,
    passed,
    detail: passed
      ? `${assertion.id} 已满足样式要求`
      : `${assertion.id} ${cell ? `样式不符合：${failedChecks.join('、')}` : '不存在或没有可检查的样式'}`,
  }
}

function hasStyleAssertion(assertion: WorkbookAssertion): boolean {
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
  return value === undefined ? '缺失' : JSON.stringify(value)
}

async function loadStyleCells(path: string): Promise<Map<string, ExcelJS.Cell>> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(stripPivotTableParts(await readFile(path)) as any)
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
