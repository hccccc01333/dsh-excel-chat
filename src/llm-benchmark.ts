import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runAgentTask, type AgentPlanner } from './agent.ts'
import { evaluateTaskChecks, type FileBenchmarkTask } from './file-benchmark.ts'
import { validateWorkbookFile } from './workbook.ts'

export interface LlmTaskResult {
  id: string
  category: string
  name: string
  success: boolean
  checksPassed: number
  checksTotal: number
  integrity: number
  rounds: number
  achieved: boolean
  error: string | null
}

export interface LlmBenchmarkReport {
  total: number
  success: number
  successRate: number
  meanAccuracy: number
  integrityRate: number
  categories: Record<string, { total: number; success: number; successRate: number }>
  tasks: LlmTaskResult[]
}

/**
 * LLM planner benchmark: run the goal-mode agent loop against every corpus
 * task with a real planner, then score the final workbook with the same
 * checks and integrity rules as the offline ExcelBench runner.
 */
export async function runLlmTask(
  task: FileBenchmarkTask,
  options: { planner: AgentPlanner; maxRounds?: number },
): Promise<LlmTaskResult> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-llm-bench-'))
  try {
    const inputPath = await task.buildInput(dir)
    const outPath = join(dir, 'out.xlsx')
    const agent = await runAgentTask(inputPath, {
      goal: task.description,
      planner: options.planner,
      maxRounds: options.maxRounds ?? 2,
      outPath,
    })
    const { passed, total } = await evaluateTaskChecks(task, agent.outputPath)
    const integrity = (await validateWorkbookFile(agent.outputPath)).anomalies.length
    return {
      id: task.id,
      category: task.category,
      name: task.name,
      success: passed === total && integrity === 0,
      checksPassed: passed,
      checksTotal: total,
      integrity,
      rounds: agent.rounds.length,
      achieved: agent.achieved,
      error: null,
    }
  } catch (error) {
    return {
      id: task.id,
      category: task.category,
      name: task.name,
      success: false,
      checksPassed: 0,
      checksTotal: task.checks.length,
      integrity: -1,
      rounds: 0,
      achieved: false,
      error: error instanceof Error ? error.message : String(error),
    }
  }
}

export async function runLlmBenchmark(
  tasks: FileBenchmarkTask[],
  options: { planner: AgentPlanner; maxRounds?: number },
): Promise<LlmBenchmarkReport> {
  const results: LlmTaskResult[] = []
  for (const task of tasks) {
    results.push(await runLlmTask(task, options))
  }
  const success = results.filter((result) => result.success).length
  const accuracySum = results.reduce((sum, result) => sum + (result.checksTotal === 0 ? 1 : result.checksPassed / result.checksTotal), 0)
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
    integrityRate: results.filter((result) => result.integrity === 0).length / (results.length || 1),
    categories,
    tasks: results,
  }
}
