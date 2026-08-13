import { runBenchmark, type BenchmarkTask } from '../src/benchmark.ts'
import { deepseekLlmTextFromEnv } from '../src/deepseek.ts'

const table = { sheet: 'Sheet1', columns: { revenue: 'B', cost: 'C' } }

const tasks: BenchmarkTask[] = [
  {
    name: 'silent-offset',
    cells: { D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C3', D5: '=B5-C5' },
    oracleCells: { D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C4', D5: '=B5-C5' },
    table,
  },
  {
    name: 'range-tail',
    cells: { D2: '=SUM(B2:C2)', D3: '=SUM(B3:C3)', D4: '=SUM(B4:C3)', D5: '=SUM(B5:C5)' },
    oracleCells: { D2: '=SUM(B2:C2)', D3: '=SUM(B3:C3)', D4: '=SUM(B4:C4)', D5: '=SUM(B5:C5)' },
    table,
  },
  {
    name: 'structure-mismatch',
    cells: { D2: '=B2-C2', D3: '=SUM(B3:C3)', D4: '=B4-C4', D5: '=B5-C5' },
    oracleCells: { D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C4', D5: '=B5-C5' },
    table,
  },
  {
    name: 'hardcode-break',
    cells: { D2: '=B2-C2', D3: '=B3-C3', D4: '100', D5: '=B5-C5' },
    oracleCells: { D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C4', D5: '=B5-C5' },
    table,
  },
  {
    name: 'clean-noop',
    cells: { D2: '=B2-C2', D3: '=B3-C3' },
    oracleCells: { D2: '=B2-C2', D3: '=B3-C3' },
    table,
  },
]

const useLlm = process.env.VERA_BENCH_LLM === '1' && Boolean(process.env.DEEPSEEK_API_KEY)
const report = await runBenchmark(tasks, {
  llm: useLlm ? deepseekLlmTextFromEnv() : undefined,
})

console.log('useLlm:', useLlm)
console.log('model:', process.env.DEEPSEEK_MODEL ?? 'deepseek-chat')
for (const task of report.tasks) {
  const route = task.llm?.score ?? task.deterministic.score
  console.log(
    `${task.task.padEnd(20)} pass=${task.passAt1 ? 'yes' : 'no '} ` +
    `det=${task.deterministic.score.accuracy.toFixed(2)} ` +
    `llm=${task.llm ? task.llm.score.accuracy.toFixed(2) : 'n/a'} ` +
    `mismatch=${route.mismatched}`,
  )
}
console.log(`Pass@1: ${report.passAt1}/${report.total}, meanAccuracy: ${report.meanAccuracy.toFixed(3)}`)
