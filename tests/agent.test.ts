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
