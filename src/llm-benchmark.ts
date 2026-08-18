import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runAgentTask, type AgentPlanner, type AgentTaskResult } from './agent.ts'
import { evaluateTaskChecks, type FileBenchmarkTask } from './file-benchmark.ts'
import {
  classifyFailure,
  summarizeFailureBreakdown,
  type FailureCategory,
  type FailureClassification,
} from './failure-taxonomy.ts'
import type { ExcelOperation } from './operations.ts'
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
  /** Failure classification for failed tasks; null for successful ones. */
  failure: FailureClassification | null
}

export interface LlmBenchmarkReport {
  total: number
  success: number
  successRate: number
  meanAccuracy: number
  integrityRate: number
  categories: Record<string, { total: number; success: number; successRate: number }>
  failureBreakdown: Record<FailureCategory, number>
  tasks: LlmTaskResult[]
}

/** Flatten every operation the agent actually executed across all rounds. */
function flattenExecutedOps(agent: AgentTaskResult): ExcelOperation[] {
  return agent.rounds.flatMap((round) => round.plan.flatMap((step) => step.operations))
}

/** Shallow-compare canonical ops against executed ops of the same name. */
function diffOperationArgs(expected: ExcelOperation[], executed: ExcelOperation[]): string[] {
  const diffs: string[] = []
  const used = new Set<number>()
  for (const wanted of expected) {
    const index = executed.findIndex((operation, i) => operation.op === wanted.op && !used.has(i))
    if (index < 0) continue
    used.add(index)
    const actual = executed[index]!
    for (const key of Object.keys(wanted)) {
      if (key === 'op') continue
      const expectedValue = JSON.stringify((wanted as Record<string, unknown>)[key])
      const actualValue = JSON.stringify((actual as Record<string, unknown>)[key])
      if (expectedValue !== actualValue) {
        diffs.push(`${wanted.op}.${key}：期望 ${expectedValue}，实际 ${actualValue}`)
      }
    }
  }
  return diffs.slice(0, 6)
}

/**
 * LLM planner benchmark: run the goal-mode agent loop against every corpus
 * task with a real planner, then score the final workbook with the same
 * checks and integrity rules as the offline ExcelBench runner. Failed tasks
 * get a deterministic failure classification (v0.35 taxonomy).
 */
export async function runLlmTask(
  task: FileBenchmarkTask,
  options: { planner: AgentPlanner; maxRounds?: number },
): Promise<LlmTaskResult> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-llm-bench-'))
  const maxRounds = options.maxRounds ?? 2
  const expectedOps = task.operations.map((operation) => operation.op)
  try {
    const inputPath = await task.buildInput(dir)
    const outPath = join(dir, 'out.xlsx')
    const agent = await runAgentTask(inputPath, {
      goal: task.description,
      planner: options.planner,
      maxRounds,
      outPath,
      deterministicAssertions: task.checks,
    })
    const { passed, total } = await evaluateTaskChecks(task, agent.outputPath)
    const integrity = (await validateWorkbookFile(agent.outputPath)).anomalies.length
    const success = passed === total && integrity === 0
    const executed = flattenExecutedOps(agent)
    const failure = success
      ? null
      : classifyFailure({
          crashed: false,
          error: null,
          verifierFalsePositive: agent.achieved,
          rounds: agent.rounds.length,
          maxRounds,
          executedOps: executed.map((operation) => operation.op),
          expectedOps,
          argDiffs: diffOperationArgs(task.operations, executed),
          checksPassed: passed,
          checksTotal: total,
          integrity,
        })
    return {
      id: task.id,
      category: task.category,
      name: task.name,
      success,
      checksPassed: passed,
      checksTotal: total,
      integrity,
      rounds: agent.rounds.length,
      achieved: agent.achieved,
      error: null,
      failure,
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
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
      error: message,
      failure: classifyFailure({
        crashed: true,
        error: message,
        verifierFalsePositive: false,
        rounds: 0,
        maxRounds,
        executedOps: [],
        expectedOps,
        argDiffs: [],
        checksPassed: 0,
        checksTotal: task.checks.length,
        integrity: -1,
      }),
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
    failureBreakdown: summarizeFailureBreakdown(
      results.filter((result) => !result.success && result.failure !== null).map((result) => result.failure!),
    ),
    tasks: results,
  }
}
