import { test } from 'node:test'
import assert from 'node:assert/strict'
import { runBenchmark, runBenchmarkTask, type BenchmarkTask } from '../src/benchmark.ts'
import { benchmarkTasks } from '../src/benchmark-cases.ts'
import type { LlmText } from '../src/advisor.ts'

const table = { sheet: 'Sheet1', columns: { revenue: 'B', cost: 'C' } }

const tasks = benchmarkTasks.filter((task) => task.name !== 'aggregate-mismatch')

/** Repairs every anomaly with the revenue-minus-cost IR the compiler expects. */
const fakeLlm: LlmText = async (prompt) => {
  const body = prompt.split('The validator found these anomalies:')[1] ?? ''
  const patternEnd = body.indexOf('\nThe column patterns are')
  const tableEnd = body.indexOf('\nThe table schema')
  const json = body.slice(0, patternEnd >= 0 ? patternEnd : tableEnd).trim()
  const anomalies = JSON.parse(json) as Array<{ cell: string }>
  return JSON.stringify({
    repairs: anomalies.map((anomaly) => ({
      id: anomaly.cell,
      baseCell: anomaly.cell,
      ir: {
        operation: 'binary',
        left: { kind: 'column', column: 'revenue' },
        right: { kind: 'column', column: 'cost' },
        operator: '-',
      },
    })),
  })
}

test('deterministic repair passes silent-offset and range-tail tasks', async () => {
  const report = await runBenchmark(tasks, {})
  assert.equal(report.total, tasks.length)
  assert.equal(report.passAt1, 8) // all pattern/fill tasks; LLM-only tasks stay red
  const structure = report.tasks.find((task) => task.task === 'structure-mismatch')!
  assert.equal(structure.deterministic.score.passes, false)
  assert.equal(structure.llm, null)
})

test('LLM route lifts structure-mismatch and hardcode-break to Pass@1', async () => {
  const report = await runBenchmark(tasks, { llm: fakeLlm })
  assert.equal(report.passAt1, tasks.length)
  const hardcode = report.tasks.find((task) => task.task === 'hardcode-break')!
  assert.equal(hardcode.deterministic.score.passes, false)
  assert.equal(hardcode.llm!.score.passes, true)
  assert.deepEqual(hardcode.llm!.llmRepairs, [{
    id: 'D4',
    kind: 'formula',
    oldValue: '100',
    newValue: '=B4-C4',
  }])
})

test('aggregate IR repair reaches Pass@1 on the aggregate-mismatch case', async () => {
  const aggregateLlm: LlmText = async () => JSON.stringify({
    repairs: [{
      id: 'D4',
      baseCell: 'D4',
      ir: { operation: 'aggregate', metric: 'revenue', function: 'SUM', filters: [] },
    }],
  })
  const result = await runBenchmarkTask(
    benchmarkTasks.find((task) => task.name === 'aggregate-mismatch')!,
    { llm: aggregateLlm },
  )
  assert.equal(result.passAt1, true)
  assert.deepEqual(result.llm!.llmRepairs, [{
    id: 'D4',
    kind: 'formula',
    oldValue: '=B4-C4',
    newValue: '=SUM(Sheet1!$B:$B)',
  }])
})

test('LLM repairs overlapping deterministic repairs are deduplicated', async () => {
  const result = await runBenchmarkTask(tasks[0]!, { llm: fakeLlm })
  assert.equal(result.passAt1, true)
  assert.equal(result.deterministic.repairs.length, 1)
  assert.equal(result.llm!.llmRepairs.length, 0)
})

test('a task without a table schema cannot use the LLM route', async () => {
  const task: BenchmarkTask = { name: 'no-table', cells: {}, oracleCells: {} }
  await assert.rejects(
    runBenchmarkTask(task, { llm: fakeLlm }),
    /needs a table schema/,
  )
})

test('an LLM failure is recorded instead of crashing the benchmark', async () => {
  const failingLlm: LlmText = async () => {
    throw new Error('model timeout')
  }
  const result = await runBenchmarkTask(
    benchmarkTasks.find((task) => task.name === 'structure-mismatch')!,
    { llm: failingLlm },
  )
  assert.equal(result.passAt1, false)
  assert.equal(result.llm, null)
  assert.equal(result.llmError, 'model timeout')
})
