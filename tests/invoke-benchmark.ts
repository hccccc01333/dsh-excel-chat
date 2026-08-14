import { runBenchmark } from '../src/benchmark.ts'
import { benchmarkTasks } from '../src/benchmark-cases.ts'
import { deepseekLlmTextFromEnv } from '../src/deepseek.ts'

const useLlm = process.env.VERA_BENCH_LLM === '1' && Boolean(process.env.DEEPSEEK_API_KEY)
const report = await runBenchmark(benchmarkTasks, {
  llm: useLlm ? deepseekLlmTextFromEnv() : undefined,
})

console.log('useLlm:', useLlm)
console.log('model:', process.env.DEEPSEEK_MODEL ?? 'deepseek-chat')
for (const task of report.tasks) {
  const route = task.llm?.score ?? task.deterministic.score
  console.log(
    `${task.task.padEnd(20)} pass=${task.passAt1 ? 'yes' : 'no '} ` +
    `det=${task.deterministic.score.accuracy.toFixed(2)} ` +
    `llm=${task.llm ? task.llm.score.accuracy.toFixed(2) : task.llmError ? 'err' : 'n/a'} ` +
    `mismatch=${route.mismatched}`,
  )
  if (task.llmError) console.log(`  llmError: ${task.llmError}`)
}
console.log(`Pass@1: ${report.passAt1}/${report.total}, meanAccuracy: ${report.meanAccuracy.toFixed(3)}`)
