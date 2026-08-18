import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { readFile, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runAgentTask, type AgentPlanner, type PlanStep } from '../src/agent.ts'
import { verifyWorkbookAssertions, type WorkbookAssertion } from '../src/verifier.ts'

interface ReplayFixture {
  goal: string
  plan: PlanStep[]
  assertions: WorkbookAssertion[]
}

async function makeWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-verifier-'))
  const path = join(dir, 'book.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '区域'
  sheet.getCell('B1').value = '金额'
  sheet.getCell('A2').value = '华东'
  sheet.getCell('B2').value = 100
  sheet.getCell('A3').value = '华北'
  sheet.getCell('B3').value = null
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

async function readReplay(): Promise<ReplayFixture> {
  return JSON.parse(await readFile(new URL('./verifier-replay.json', import.meta.url), 'utf8')) as ReplayFixture
}

test('deterministic verifier returns evidence and failures without an LLM', async () => {
  const path = await makeWorkbook()
  const failed = await verifyWorkbookAssertions(path, [
    { id: '订单!B3', expect: '0' },
    { id: '订单!A1', bold: true },
  ])
  assert.equal(failed.achieved, false)
  assert.equal(failed.passed, 0)
  assert.equal(failed.total, 2)
  assert.match(failed.reason, /订单!B3/)
  assert.match(failed.reason, /订单!A1/)
})

test('replay assertions hard-gate a fake verifier after a correct plan', async () => {
  const replay = await readReplay()
  const path = await makeWorkbook()
  let verifyCalls = 0
  const planner: AgentPlanner = {
    async plan() {
      return replay.plan
    },
    async verify() {
      verifyCalls += 1
      return { achieved: false, reason: 'fake verifier should not run' }
    },
  }
  const result = await runAgentTask(path, {
    goal: replay.goal,
    planner,
    deterministicAssertions: replay.assertions,
  })
  assert.equal(result.achieved, true)
  assert.equal(verifyCalls, 0)
  assert.deepEqual(result.rounds[0]!.deterministicVerification?.failures, [])
})

test('deterministic assertion failure overrides a false-positive fake verifier', async () => {
  const path = await makeWorkbook()
  let verifyCalls = 0
  const planner: AgentPlanner = {
    async plan() {
      return [{ name: 'wrong edit', operations: [{ op: 'set', cells: { '订单!A2': '华南' } }] }]
    },
    async verify() {
      verifyCalls += 1
      return { achieved: true, reason: 'fake verifier incorrectly says complete' }
    },
  }
  const result = await runAgentTask(path, {
    goal: '把金额空值补为 0',
    planner,
    maxRounds: 1,
    deterministicAssertions: [{ id: '订单!B3', expect: '0' }],
  })
  assert.equal(result.achieved, false)
  assert.equal(verifyCalls, 0)
  assert.match(result.rounds[0]!.verdict.reason, /订单!B3/)
})
