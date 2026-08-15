import { readFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import ExcelJS from 'exceljs'
import { autofixWorkbookFile } from './autofix.ts'
import { applyOperationsToWorkbook, type ExcelOperation } from './operations.ts'
import { validate } from './validator.ts'
import { normalizeCellId } from './score.ts'
import { readWorkbookCells, stripPivotTableParts } from './workbook.ts'

export type FileBenchmarkCategory = 'formula' | 'editing' | 'analysis' | 'workflow'

export interface FileCheck {
  /** Cell id in the output workbook, e.g. "订单!D4". */
  id: string
  /** Expected exact cell content; null means the cell must be absent. */
  expect?: string | null
  /** Expected cell-content prefix (for formulas like "=SUMIFS("). */
  startsWith?: string
  /** Style assertions, evaluated by loading the output workbook. */
  fill?: string
  bold?: boolean
  numberFormat?: string
  wrapText?: boolean
  hAlign?: string
}

export interface FileBenchmarkTask {
  id: string
  category: FileBenchmarkCategory
  name: string
  description: string
  /** Build the input workbook inside dir and return its absolute path. */
  buildInput: (dir: string) => Promise<string>
  /** The canonical operation plan the agent is expected to execute. */
  operations: ExcelOperation[]
  /** Assertions on the final workbook. */
  checks: FileCheck[]
  /** Evaluate checks after the verification/autofix pass (for repair tasks). */
  evaluateAfterAutofix?: boolean
}

export interface FileTaskResult {
  id: string
  category: FileBenchmarkCategory
  name: string
  checksPassed: number
  checksTotal: number
  success: boolean
  accuracy: number
  integrityBefore: number
  integrityAfter: number
  repaired: number
}

export interface FileBenchmarkReport {
  total: number
  success: number
  successRate: number
  meanAccuracy: number
  integrityRate: number
  categories: Record<string, { total: number; success: number; successRate: number }>
  tasks: FileTaskResult[]
}

/**
 * Offline, file-based benchmark (ExcelBench lite): each realistic task builds
 * an input workbook, executes the canonical operation plan, then asserts
 * cell/style outcomes and workbook integrity (formula anomalies after
 * verification+repair). Success = all checks pass AND the workbook is clean
 * after the verify/repair pass.
 */
export async function runFileBenchmarkTask(
  task: FileBenchmarkTask,
  dir: string,
): Promise<FileTaskResult> {
  const inputPath = await task.buildInput(dir)
  const outputPath = join(dir, `${task.id}.out.xlsx`)
  await applyOperationsToWorkbook(inputPath, task.operations, outputPath)

  let workingPath = outputPath
  let cells = await readWorkbookCells(await readFile(outputPath))
  const integrityBefore = validate(cells).anomalies.length
  let integrityAfter = integrityBefore
  let repaired = 0
  if (integrityBefore > 0) {
    const fix = await autofixWorkbookFile(outputPath)
    repaired = fix.repairs.length
    integrityAfter = fix.after.total
    workingPath = fix.repairedPath
    if (task.evaluateAfterAutofix) {
      cells = await readWorkbookCells(await readFile(workingPath))
    }
  }

  const styleChecks = task.checks.filter((check) => check.fill !== undefined || check.bold !== undefined || check.numberFormat !== undefined || check.wrapText !== undefined || check.hAlign !== undefined)
  const checksPath = task.evaluateAfterAutofix ? workingPath : outputPath
  const styleCells = styleChecks.length > 0 ? await loadStyleCells(checksPath) : null
  let checksPassed = 0
  for (const check of task.checks) {
    const normalized = normalizeCellId(check.id)
    const actual = cells[normalized] ?? cells[findKey(cells, normalized) ?? '']
    if (check.expect !== undefined) {
      const exists = actual !== undefined && actual !== ''
      if (check.expect === null ? !exists : actual === check.expect) checksPassed += 1
      continue
    }
    if (check.startsWith !== undefined) {
      if (typeof actual === 'string' && actual.startsWith(check.startsWith)) checksPassed += 1
      continue
    }
    const cell = styleCells?.get(normalized)
    if (!cell) continue
    if (check.fill !== undefined) {
      const argb = cell.fill?.type === 'pattern' ? (cell.fill.fgColor as { argb?: string } | undefined)?.argb?.toUpperCase() : undefined
      if (argb?.endsWith(check.fill.toUpperCase())) checksPassed += 1
    }
    if (check.bold !== undefined && (cell.font?.bold ?? false) === check.bold) checksPassed += 1
    if (check.numberFormat !== undefined && cell.numFmt === check.numberFormat) checksPassed += 1
    if (check.wrapText !== undefined && (cell.alignment?.wrapText ?? false) === check.wrapText) checksPassed += 1
    if (check.hAlign !== undefined && cell.alignment?.horizontal === check.hAlign) checksPassed += 1
  }

  const checksTotal = task.checks.length
  return {
    id: task.id,
    category: task.category,
    name: task.name,
    checksPassed,
    checksTotal,
    success: checksPassed === checksTotal && integrityAfter === 0,
    accuracy: checksTotal === 0 ? 1 : checksPassed / checksTotal,
    integrityBefore,
    integrityAfter,
    repaired,
  }
}

export async function runFileBenchmark(tasks: FileBenchmarkTask[]): Promise<FileBenchmarkReport> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-bench-'))
  const results: FileTaskResult[] = []
  for (const task of tasks) {
    results.push(await runFileBenchmarkTask(task, dir))
  }
  const success = results.filter((result) => result.success).length
  const accuracySum = results.reduce((sum, result) => sum + result.accuracy, 0)
  const categories: Record<string, { total: number; success: number; successRate: number }> = {}
  for (const result of results) {
    const entry = (categories[result.category] ??= { total: 0, success: 0, successRate: 0 })
    entry.total += 1
    if (result.success) entry.success += 1
  }
  for (const entry of Object.values(categories)) entry.successRate = entry.total === 0 ? 0 : entry.success / entry.total
  return {
    total: results.length,
    success,
    successRate: results.length === 0 ? 0 : success / results.length,
    meanAccuracy: results.length === 0 ? 0 : accuracySum / results.length,
    integrityRate: results.filter((result) => result.integrityAfter === 0).length / (results.length || 1),
    categories,
    tasks: results,
  }
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
