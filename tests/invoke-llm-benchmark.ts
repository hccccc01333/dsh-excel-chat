/**
 * Real LLM planner benchmark: runs the goal-mode agent loop (DeepSeek
 * planner + verifier) against the corpus and prints the report.
 *   node --test tests/invoke-llm-benchmark.ts
 * Env: DEEPSEEK_API_KEY required; LLM_BENCH_SAMPLE limits to N tasks,
 * LLM_BENCH_OFFSET skips the first M tasks (default 0).
 */
import { corpusTasks } from '../src/corpus/index.ts'
import { deepseekLlmTextFromEnv } from '../src/deepseek.ts'
import { runLlmBenchmark } from '../src/llm-benchmark.ts'
import { createLlmPlanner } from '../src/llm-planner.ts'

const model = process.env.LLM_PROVIDER === 'bai'
  ? (process.env.BAI_MODEL ?? 'glm-5.3-flash')
  : (process.env.DEEPSEEK_MODEL ?? 'deepseek-chat')
const sample = process.env.LLM_BENCH_SAMPLE ? Number(process.env.LLM_BENCH_SAMPLE) : corpusTasks.length
const offset = process.env.LLM_BENCH_OFFSET ? Number(process.env.LLM_BENCH_OFFSET) : 0
const tasks = corpusTasks.slice(offset, offset + sample)
const planner = createLlmPlanner(deepseekLlmTextFromEnv(model))
const report = await runLlmBenchmark(tasks, { planner, maxRounds: 3 })
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
