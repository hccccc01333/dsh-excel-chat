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

/**
 * An assertion that names several conditions is a conjunction. The verifier used
 * to return after the first kind of condition matched, so `{ expect, bold }`
 * checked the value and ignored the style — reporting "achieved" for a workbook
 * whose formatting was never applied. That is the failure mode this verifier
 * exists to prevent, so each combination is pinned here.
 */
test('an assertion must satisfy every declared condition, not just the first', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-verifier-conj-'))
  const path = join(dir, 'styled.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '区域'
  sheet.getCell('B1').value = '金额'
  sheet.getCell('B1').font = { bold: true }
  sheet.getCell('A2').value = '华东'
  sheet.getCell('B2').value = 100
  sheet.getCell('B2').numFmt = '#,##0.00'
  sheet.getCell('B2').alignment = { horizontal: 'right' }
  sheet.getCell('B2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }
  // A3 holds the right value but none of the styling the assertions below ask for.
  sheet.getCell('A3').value = '华北'
  await writeFile(path, await workbook.xlsx.writeBuffer())

  // Every condition holds on B2.
  const all = await verifyWorkbookAssertions(path, [
    { id: '订单!B2', expect: '100', numberFormat: '#,##0.00', hAlign: 'right', fill: 'FFD9D9D9' },
  ])
  assert.equal(all.achieved, true, all.reason)

  // The value matches but the style does not: must fail, and name the condition.
  const styleMiss = await verifyWorkbookAssertions(path, [
    { id: '订单!A3', expect: '华北', bold: true },
  ])
  assert.equal(styleMiss.achieved, false, 'a matching value must not excuse a missing style')
  assert.match(styleMiss.failures[0]!, /加粗/)
  assert.doesNotMatch(styleMiss.failures[0]!, /值 期望/, 'the condition that did hold must not be listed')

  // The style matches but the value does not: must also fail.
  const valueMiss = await verifyWorkbookAssertions(path, [
    { id: '订单!B2', expect: '999', numberFormat: '#,##0.00' },
  ])
  assert.equal(valueMiss.achieved, false)
  assert.match(valueMiss.failures[0]!, /值 期望/)
  assert.doesNotMatch(valueMiss.failures[0]!, /数字格式/, 'the style that did hold must not be listed')

  // A prefix condition alongside a style is checked too.
  const prefixMiss = await verifyWorkbookAssertions(path, [
    { id: '订单!B2', startsWith: '=', bold: true },
  ])
  assert.equal(prefixMiss.achieved, false)
  assert.match(prefixMiss.failures[0]!, /前缀/)
  assert.match(prefixMiss.failures[0]!, /加粗/)
})

test('a passing assertion reports how many conditions it checked', async () => {
  const path = await makeWorkbook()
  const result = await verifyWorkbookAssertions(path, [{ id: '订单!B2', expect: '100', bold: false }])
  assert.equal(result.achieved, true, result.reason)
  assert.match(result.assertions[0]!.detail, /2 项要求/)
})
