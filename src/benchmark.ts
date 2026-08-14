import { createLlmRepairAdvisor, type LlmText } from './advisor.ts'
import type { ColumnTable } from './ir.ts'
import { applyPatches, type CellPatch } from './patch.ts'
import { generateRepairs } from './repair.ts'
import { scoreWorkbookAgainstOracle, type WorkbookScore } from './score.ts'
import { validate } from './validator.ts'

export interface BenchmarkTask {
  name: string
  /** Starting workbook cells the agent is asked to repair. */
  cells: Record<string, string>
  /** Ground-truth workbook the agent should reproduce. */
  oracleCells: Record<string, string>
  /** Table schema required for LLM repair; optional for deterministic-only runs. */
  table?: ColumnTable
}

export interface RouteResult {
  repairs: CellPatch[]
  llmRepairs: CellPatch[]
  finalCells: Record<string, string>
  score: WorkbookScore
}

export interface BenchmarkTaskResult {
  task: string
  deterministic: RouteResult
  llm: RouteResult | null
  llmError: string | null
  passAt1: boolean
}

export interface BenchmarkReport {
  generatedAt: string
  tasks: BenchmarkTaskResult[]
  passAt1: number
  total: number
  meanAccuracy: number
}

export interface BenchmarkOptions {
  llm?: LlmText
  signal?: AbortSignal
}

/**
 * Run one task the way the repair tool does in production: deterministic
 * pattern repairs first, then an optional LLM advisor on the original
 * workbook + validation result for anomalies deterministic repair cannot
 * cover. Pass@1 means the final workbook matches the oracle cell for cell.
 */
export async function runBenchmarkTask(
  task: BenchmarkTask,
  options: BenchmarkOptions = {},
): Promise<BenchmarkTaskResult> {
  const before = validate(task.cells)
  const repairs = generateRepairs(task.cells, before)
  const deterministicCells = repairs.length > 0 ? applyPatches(task.cells, repairs) : task.cells
  const deterministic: RouteResult = {
    repairs,
    llmRepairs: [],
    finalCells: deterministicCells,
    score: scoreWorkbookAgainstOracle(deterministicCells, task.oracleCells),
  }

  let llm: RouteResult | null = null
  let llmError: string | null = null
  if (options.llm) {
    if (!task.table) {
      throw new Error(`task "${task.name}" needs a table schema for LLM repair`)
    }
    const advisor = createLlmRepairAdvisor(options.llm, task.table, options.signal)
    try {
      const llmRepairs = await advisor(task.cells, before)
      const covered = new Set(repairs.map((patch) => patch.id))
      const extraLlmRepairs = llmRepairs.filter((patch) => !covered.has(patch.id))
      const llmCells = extraLlmRepairs.length > 0 ? applyPatches(deterministicCells, extraLlmRepairs) : deterministicCells
      llm = {
        repairs,
        llmRepairs: extraLlmRepairs,
        finalCells: llmCells,
        score: scoreWorkbookAgainstOracle(llmCells, task.oracleCells),
      }
    } catch (error) {
      llmError = error instanceof Error ? error.message : String(error)
    }
  }

  return {
    task: task.name,
    deterministic,
    llm,
    llmError,
    passAt1: deterministic.score.passes || (llm?.score.passes ?? false),
  }
}

export async function runBenchmark(
  tasks: BenchmarkTask[],
  options: BenchmarkOptions = {},
): Promise<BenchmarkReport> {
  const results: BenchmarkTaskResult[] = []
  for (const task of tasks) {
    results.push(await runBenchmarkTask(task, options))
  }
  const passed = results.filter((result) => result.passAt1).length
  const accuracySum = results.reduce(
    (sum, result) => sum + (result.llm?.score.accuracy ?? result.deterministic.score.accuracy),
    0,
  )
  return {
    generatedAt: new Date().toISOString(),
    tasks: results,
    passAt1: passed,
    total: results.length,
    meanAccuracy: results.length === 0 ? 0 : accuracySum / results.length,
  }
}
