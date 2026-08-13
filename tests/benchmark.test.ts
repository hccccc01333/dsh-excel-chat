import { test } from 'node:test'
import assert from 'node:assert/strict'
import { runBenchmark, runBenchmarkTask, type BenchmarkTask } from '../src/benchmark.ts'
import type { LlmText } from '../src/advisor.ts'

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

/** Repairs every anomaly with the revenue-minus-cost IR the compiler expects. */
const fakeLlm: LlmText = async (prompt) => {
  const body = prompt.split('The validator found these anomalies:')[1] ?? ''
  const json = body.slice(0, body.indexOf('\nThe table schema'))
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
  assert.equal(report.passAt1, 3) // silent-offset, range-tail, clean-noop
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
