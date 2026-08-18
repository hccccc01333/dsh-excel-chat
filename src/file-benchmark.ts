import { readFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { autofixWorkbookFile } from './autofix.ts'
import { applyOperationsToWorkbook, type ExcelOperation } from './operations.ts'
import { validate } from './validator.ts'
import { verifyWorkbookAssertions, type WorkbookAssertion } from './verifier.ts'
import { readWorkbookCells } from './workbook.ts'

export type FileBenchmarkCategory = 'formula' | 'editing' | 'analysis' | 'workflow'

export type FileCheck = WorkbookAssertion

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

export interface CheckResult {
  passed: number
  total: number
}

/** Evaluate a task's checks against an output workbook (shared by offline and LLM runners). */
export async function evaluateTaskChecks(task: FileBenchmarkTask, outputPath: string): Promise<CheckResult> {
  const verification = await verifyWorkbookAssertions(outputPath, task.checks)
  return { passed: verification.passed, total: verification.total }
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

  const checksPath = task.evaluateAfterAutofix ? workingPath : outputPath
  const { passed: checksPassed, total: checksTotal } = await evaluateTaskChecks(task, checksPath)
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
