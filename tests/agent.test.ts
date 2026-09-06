import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { access, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runAgentTask, type AgentPlanContext, type AgentPlanner, type PlanStep } from '../src/agent.ts'

async function makeWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-agent-'))
  const path = join(dir, 'book.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '区域'
  sheet.getCell('B1').value = '金额'
  sheet.getCell('A2').value = '华东'
  sheet.getCell('B2').value = 100
  sheet.getCell('A3').value = '华东'
  sheet.getCell('B3').value = null
  sheet.getCell('A4').value = '华北'
  sheet.getCell('B4').value = 300
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

function fakePlanner(
  plans: PlanStep[][],
  verdicts: Array<{ achieved: boolean; reason: string }>,
): AgentPlanner & { seen: AgentPlanContext[] } {
  const seen: AgentPlanContext[] = []
  return {
    seen,
    async plan(context) {
      seen.push(context)
      return plans[Math.min(context.round - 1, plans.length - 1)]!
    },
    async verify(_context) {
      return verdicts[Math.min(_context.round - 1, verdicts.length - 1)]!
    },
  }
}

test('agent loop achieves the goal in one round and writes the final file', async () => {
  const path = await makeWorkbook()
  const planner = fakePlanner(
    [[{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }]],
    [{ achieved: true, reason: '完成' }],
  )
  const result = await runAgentTask(path, { goal: '把金额空值填 0', planner, maxRounds: 2 })
  assert.equal(result.rounds.length, 1)
  assert.equal(result.achieved, true)
  assert.equal(result.finalAnomalies, 0)
  assert.match(result.outputPath, /\.agent\.xlsx$/)
  await access(result.outputPath)
})

test('agent loop replans when the verifier says the goal is not achieved', async () => {
  const path = await makeWorkbook()
  const planner = fakePlanner(
    [
      [{ name: 'noop', operations: [{ op: 'set', cells: { '订单!E1': 'x' } }] }],
      [{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }],
    ],
    [
      { achieved: false, reason: '还差补空值' },
      { achieved: true, reason: '完成' },
    ],
  )
  const result = await runAgentTask(path, { goal: '补空值', planner, maxRounds: 3 })
  assert.equal(result.rounds.length, 2)
  assert.equal(result.achieved, true)
  const second = planner.seen[1]!
  assert.equal(second.round, 2)
  assert.ok(second.previousPlan)
  assert.equal(second.verifierNote, '还差补空值')
})

test('agent loop stops at maxRounds when the goal is never achieved', async () => {
  const path = await makeWorkbook()
  const planner = fakePlanner(
    [
      [{ name: 'a', operations: [{ op: 'set', cells: { '订单!E1': 'x' } }] }],
      [{ name: 'b', operations: [{ op: 'set', cells: { '订单!E2': 'y' } }] }],
    ],
    [
      { achieved: false, reason: '未完成' },
      { achieved: false, reason: '仍未完成' },
    ],
  )
  const result = await runAgentTask(path, { goal: '不可能的任务', planner, maxRounds: 2 })
  assert.equal(result.rounds.length, 2)
  assert.equal(result.achieved, false)
})

test('agent loop feeds an invalid plan back to the planner for a corrected round', async () => {
  const path = await makeWorkbook()
  const planner = fakePlanner(
    [
      [{ name: 'bad', operations: [{ op: 'aggregateReport', source: '订单!A1:B4', groupColumn: 'A' }] }],
      [{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }],
    ],
    [
      { achieved: true, reason: 'ok' },
      { achieved: true, reason: '完成' },
    ],
  )
  const result = await runAgentTask(path, { goal: '补空值', planner, maxRounds: 2 })
  assert.equal(result.rounds.length, 1)
  assert.equal(result.achieved, true)
  assert.ok(planner.seen[1]!.verifierNote?.includes('计划无效'))
})

test('Verifier 2.0: failing plan assertions block the verdict and feed back', async () => {
  const path = await makeWorkbook()
  const seen: AgentPlanContext[] = []
  let calls = 0
  const planner: AgentPlanner = {
    async plan(context) {
      calls++
      seen.push(context)
      if (calls === 1) {
        // Plan fills the blanks but asserts a wrong value on purpose.
        return {
          steps: [{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }],
          assertions: [{ id: '订单!B3', expect: 999 }],
        }
      }
      return {
        steps: [{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }, { name: 'label', operations: [{ op: 'set', cells: { '订单!E1': '已补齐' } }] }],
        assertions: [{ id: '订单!B3', expect: 0 }, { id: '订单!E1', expect: '已补齐' }],
      }
    },
    async verify() {
      return { achieved: true, reason: 'LLM 认为完成' }
    },
  }
  const result = await runAgentTask(path, { goal: '补空值', planner, maxRounds: 3 })
  // Round 1: LLM says achieved, but the planner's own assertion fails -> replan.
  assert.ok(result.rounds[0]!.planAssertions)
  assert.equal(result.rounds[0]!.planAssertions!.achieved, false)
  assert.equal(result.rounds[0]!.verdict.achieved, false)
  assert.match(result.rounds[0]!.verdict.reason, /规划器断言未过/)
  assert.ok(seen[1]!.verifierNote?.includes('规划器断言未过'))
  // Round 2: assertions pass -> goal achieved despite the LLM being noisy-positive.
  assert.equal(result.rounds.length, 2)
  assert.equal(result.achieved, true)
})

test('Verifier 2.0: passing plan assertions cannot override a negative LLM verdict', async () => {
  const path = await makeWorkbook()
  const planner: AgentPlanner = {
    async plan() {
      return {
        steps: [{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }],
        assertions: [{ id: '订单!B3', expect: 0 }],
      }
    },
    async verify() {
      return { achieved: false, reason: '还缺别的' }
    },
  }
  const result = await runAgentTask(path, { goal: '补空值', planner, maxRounds: 1 })
  assert.equal(result.rounds[0]!.planAssertions!.achieved, true)
  assert.equal(result.rounds[0]!.verdict.achieved, false)
  assert.equal(result.achieved, false)
})

test('Verifier 2.0: invalid planner assertions are dropped without breaking the loop', async () => {
  const path = await makeWorkbook()
  const planner: AgentPlanner = {
    async plan() {
      return {
        steps: [{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }],
        assertions: [{ expect: 1 }, { id: '', expect: 'x' }, { id: '订单!B3' }, { id: '订单!B3', expect: 0 }],
      }
    },
    async verify() {
      return { achieved: true, reason: '完成' }
    },
  }
  const result = await runAgentTask(path, { goal: '补空值', planner, maxRounds: 1 })
  // Only the one well-formed assertion survives sanitation and passes.
  assert.equal(result.rounds[0]!.planAssertions!.total, 1)
  assert.equal(result.rounds[0]!.planAssertions!.achieved, true)
  assert.equal(result.achieved, true)
})
