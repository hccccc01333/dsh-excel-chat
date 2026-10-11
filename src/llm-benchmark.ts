import { appendFile, mkdir, readFile } from 'node:fs/promises'
import { dirname } from 'node:path'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runAgentTask, type AgentPlanner, type AgentTaskResult } from './agent.ts'
import { t } from './i18n.ts'
import { PlanSchemaError } from './plan-schema.ts'
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
        diffs.push(t('{op}.{key}：期望 {expected}，实际 {actual}', { op: wanted.op, key, expected: expectedValue, actual: actualValue }))
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
    // Same rule as the offline runner: a task that asserts nothing must not count
    // as a success, or a corpus typo silently lifts the reported rate.
    const success = total > 0 && passed === total && integrity === 0
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
    // Carry the typed kind through: the taxonomy classifies on it rather than on
    // the message wording, which is translated.
    const errorKind = error instanceof PlanSchemaError ? error.kind : undefined
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
        errorKind,
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

export interface LlmBenchmarkOptions {
  planner: AgentPlanner
  maxRounds?: number
  /**
   * JSONL checkpoint file: every completed task is appended as one line and
   * previously recorded task ids are skipped on rerun, so sliced or crashed
   * benchmark runs can resume without losing progress. Crashed tasks (API
   * errors) are NOT checkpointed so the next invocation retries them.
   */
  outFile?: string
  /** Called after each freshly executed task (not for checkpoint-skipped ones). */
  onProgress?: (result: LlmTaskResult, completed: number, total: number) => void
  /** Retry attempts for transient provider errors (default 3). */
  retries?: number
  /** Base backoff between transient-error retries in ms (default 20000). */
  retryDelayMs?: number
  /** Pause between tasks in ms to stay under provider rate limits. */
  interTaskDelayMs?: number
  /** How many tasks run at once. Tasks are independent, so this only affects the clock. */
  concurrency?: number
}

/** Read a JSONL checkpoint file into id -> result (last line wins). */
export async function readBenchmarkCheckpoint(outFile: string): Promise<Map<string, LlmTaskResult>> {
  const completed = new Map<string, LlmTaskResult>()
  let text: string
  try {
    text = await readFile(outFile, 'utf8')
  } catch {
    return completed
  }
  for (const line of text.split('\n')) {
    if (!line.trim()) continue
    try {
      const result = JSON.parse(line) as LlmTaskResult
      if (result && typeof result.id === 'string') completed.set(result.id, result)
    } catch {
      // Tolerate a torn final line from a killed process.
    }
  }
  return completed
}

  /** Server-side failures worth retrying: rate limits, outages, network blips. */
  const TRANSIENT_ERROR = /429|500|503|fetch failed|temporarily_unavailable|query_data_error/i

  /** A rate limit is not the same kind of failure as a network blip. */
  const RATE_LIMIT = /\b429\b/

  function sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms))
  }

  /**
   * How long to wait before retrying a transient failure.
   *
   * Rate limits get a floor of a minute. The gateway sends no `Retry-After` to say how long
   * it wants, and linear backoff on a 20-second base is nowhere near enough: at concurrency
   * 8 the three attempts all landed inside the same rate-limit window, every one came back
   * 429, and twelve tasks were recorded as failures that had nothing to do with the planner.
   * A retry that returns before the window closes is not a retry.
   */
  export function retryDelayMsFor(error: string, baseMs: number, attempt: number): number {
    const base = RATE_LIMIT.test(error) ? Math.max(baseMs, 60_000) : baseMs
    return base * attempt
  }

  /**
   * Run one task, retrying transient provider failures (429/503/network) with
   * linear backoff so flaky hosted endpoints don't poison the whole report.
   */
  export async function runLlmTaskWithRetry(
    task: FileBenchmarkTask,
    options: { planner: AgentPlanner; maxRounds?: number; retries?: number; retryDelayMs?: number },
  ): Promise<LlmTaskResult> {
    const attempts = Math.max(1, options.retries ?? 3)
    const delay = options.retryDelayMs ?? 20_000
    let last = await runLlmTask(task, options)
    for (let attempt = 1; attempt < attempts && last.error !== null && TRANSIENT_ERROR.test(last.error); attempt++) {
      await sleep(retryDelayMsFor(last.error, delay, attempt))
      last = await runLlmTask(task, options)
    }
    return last
  }

export async function runLlmBenchmark(
  tasks: FileBenchmarkTask[],
  options: LlmBenchmarkOptions,
): Promise<LlmBenchmarkReport> {
  const { outFile, onProgress } = options
  const checkpoint = outFile ? await readBenchmarkCheckpoint(outFile) : new Map<string, LlmTaskResult>()
  // The tasks are independent — each builds its own input in its own scratch directory —
  // so running several at once changes nothing about the result and everything about the
  // wall clock. Serial, a full corpus took hours; the model answers in about four
  // seconds, and almost all of that time was spent waiting.
  const concurrency = Math.max(1, Math.min(options.concurrency ?? 1, tasks.length || 1))
  const results: Array<LlmTaskResult | undefined> = new Array(tasks.length)
  let completed = 0
  let next = 0

  // Appends are serialised: two workers finishing together must not interleave halves of
  // two JSON lines. The checkpoint is the raw result, so a corrupted line is a lost task.
  let appendChain: Promise<void> = Promise.resolve()
  const appendResult = (result: LlmTaskResult) => {
    if (!outFile) return
    appendChain = appendChain.then(async () => {
      await mkdir(dirname(outFile), { recursive: true })
      await appendFile(outFile, `${JSON.stringify(result)}\n`, 'utf8')
    })
    return appendChain
  }

  const worker = async () => {
    for (;;) {
      const index = next++
      if (index >= tasks.length) return
      const task = tasks[index]!
      const prior = checkpoint.get(task.id)
      if (prior) {
        results[index] = prior
        completed += 1
        onProgress?.(prior, completed, tasks.length)
        continue
      }
      const result = await runLlmTaskWithRetry(task, options)
      results[index] = result
      completed += 1
      // Every result is written, pass or fail. Keeping only `error === null` made the
      // file a record of what worked: a run that hit 144 tasks left 70 lines behind, and
      // the 74 that were missing were exactly the hard ones. The rate computed from that
      // file was 78.6% over a biased subset, and nothing said so.
      //
      // Retrying is `runLlmTaskWithRetry`'s job and it has already finished by now, so
      // what arrives here is the final word on this task either way.
      await appendResult(result)
      onProgress?.(result, completed, tasks.length)
      if (options.interTaskDelayMs) await sleep(options.interTaskDelayMs)
    }
  }

  await Promise.all(Array.from({ length: concurrency }, worker))
  await appendChain
  const finished = results.filter((result): result is LlmTaskResult => result !== undefined)
  const success = finished.filter((result) => result.success).length
  const accuracySum = finished.reduce((sum, result) => sum + (result.checksTotal === 0 ? 0 : result.checksPassed / result.checksTotal), 0)
  const categories: Record<string, { total: number; success: number; successRate: number }> = {}
  for (const result of finished) {
    const entry = (categories[result.category] ??= { total: 0, success: 0, successRate: 0 })
    entry.total += 1
    if (result.success) entry.success += 1
  }
  for (const entry of Object.values(categories)) entry.successRate = entry.total === 0 ? 0 : entry.success / entry.total
  return {
    total: finished.length,
    success,
    successRate: finished.length === 0 ? 0 : success / finished.length,
    meanAccuracy: finished.length === 0 ? 0 : accuracySum / finished.length,
    integrityRate: finished.filter((result) => result.integrity === 0).length / (finished.length || 1),
    categories,
    failureBreakdown: summarizeFailureBreakdown(
      finished.filter((result) => !result.success && result.failure !== null).map((result) => result.failure!),
    ),
    tasks: finished,
  }
}
