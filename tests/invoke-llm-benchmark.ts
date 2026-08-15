/**
 * Real LLM planner benchmark: runs the goal-mode agent loop (DeepSeek
 * planner + verifier) against the corpus and prints the report.
 *   node --test tests/invoke-llm-benchmark.ts
 * Env: DEEPSEEK_API_KEY required; LLM_BENCH_SAMPLE limits to the first N tasks.
 */
import { corpusTasks } from '../src/corpus/index.ts'
import { deepseekLlmTextFromEnv } from '../src/deepseek.ts'
import { runLlmBenchmark } from '../src/llm-benchmark.ts'
import { createLlmPlanner } from '../src/llm-planner.ts'

const model = process.env.DEEPSEEK_MODEL ?? 'deepseek-chat'
const sample = process.env.LLM_BENCH_SAMPLE ? Number(process.env.LLM_BENCH_SAMPLE) : corpusTasks.length
const tasks = corpusTasks.slice(0, sample)
const planner = createLlmPlanner(deepseekLlmTextFromEnv(model))
const report = await runLlmBenchmark(tasks, { planner, maxRounds: 2 })
console.log(JSON.stringify({
  model,
  total: report.total,
  success: report.success,
  successRate: report.successRate,
  meanAccuracy: report.meanAccuracy,
  integrityRate: report.integrityRate,
  categories: report.categories,
  failed: report.tasks.filter((task) => !task.success).map((task) => ({
    id: task.id,
    checks: `${task.checksPassed}/${task.checksTotal}`,
    integrity: task.integrity,
    rounds: task.rounds,
    achieved: task.achieved,
    error: task.error,
  })),
}, null, 2))
