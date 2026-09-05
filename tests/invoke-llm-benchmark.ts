/**
 * Real LLM planner benchmark: runs the goal-mode agent loop (planner +
 * verifier) against the corpus and prints the report.
 *   node --test tests/invoke-llm-benchmark.ts
 * Env:
 *   DEEPSEEK_API_KEY or BAI_API  — provider key (LLM_PROVIDER=bai|deepseek)
 *   BAI_MODEL / DEEPSEEK_MODEL   — model override
 *   LLM_BENCH_SAMPLE / LLM_BENCH_OFFSET — slice the corpus (default all)
 *   LLM_BENCH_OUT  — JSONL checkpoint file; completed tasks are appended and
 *                    skipped on rerun, so sliced runs can resume after a crash
 *   LLM_BENCH_ROUNDS — max replanning rounds (default 3)
 * Progress goes to stderr; the final report is the only stdout JSON.
 */
import { corpusTasks } from '../src/corpus/index.ts'
import { deepseekLlmTextFromEnv } from '../src/deepseek.ts'
import { runLlmBenchmark, type LlmTaskResult } from '../src/llm-benchmark.ts'
import { createLlmPlanner } from '../src/llm-planner.ts'

const model = process.env.LLM_PROVIDER === 'bai'
  ? (process.env.BAI_MODEL ?? 'glm-5.3-flash')
  : (process.env.DEEPSEEK_MODEL ?? 'deepseek-chat')
const sample = process.env.LLM_BENCH_SAMPLE ? Number(process.env.LLM_BENCH_SAMPLE) : corpusTasks.length
const offset = process.env.LLM_BENCH_OFFSET ? Number(process.env.LLM_BENCH_OFFSET) : 0
const maxRounds = process.env.LLM_BENCH_ROUNDS ? Number(process.env.LLM_BENCH_ROUNDS) : 3
const tasks = corpusTasks.slice(offset, offset + sample)
const outFile = process.env.LLM_BENCH_OUT || undefined
const planner = createLlmPlanner(deepseekLlmTextFromEnv(model))
const report = await runLlmBenchmark(tasks, {
  planner,
  maxRounds,
  outFile,
  retries: process.env.LLM_BENCH_RETRIES ? Number(process.env.LLM_BENCH_RETRIES) : 4,
  retryDelayMs: process.env.LLM_BENCH_RETRY_DELAY ? Number(process.env.LLM_BENCH_RETRY_DELAY) : 30_000,
  interTaskDelayMs: process.env.LLM_BENCH_TASK_DELAY ? Number(process.env.LLM_BENCH_TASK_DELAY) : 3_000,
  onProgress: (result: LlmTaskResult, completed: number, total: number) => {
    const mark = result.success ? 'PASS' : `FAIL(${result.failure?.category ?? 'crashed'})`
    console.error(`[${model}] ${completed}/${total} ${result.id} ${mark} checks=${result.checksPassed}/${result.checksTotal}`)
  },
})
console.log(JSON.stringify({
  model,
  total: report.total,
  success: report.success,
  successRate: report.successRate,
  meanAccuracy: report.meanAccuracy,
  integrityRate: report.integrityRate,
  categories: report.categories,
  failureBreakdown: report.failureBreakdown,
  failed: report.tasks.filter((task) => !task.success).map((task) => ({
    id: task.id,
    checks: `${task.checksPassed}/${task.checksTotal}`,
    integrity: task.integrity,
    rounds: task.rounds,
    achieved: task.achieved,
    error: task.error,
    failure: task.failure,
  })),
}, null, 2))
