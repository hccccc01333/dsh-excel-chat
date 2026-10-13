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

/**
 * The feedback for a second round has to be specific enough to act on.
 *
 * "not achieved" tells the planner nothing it can fix. What it needs is which
 * assertion failed and what was there instead — otherwise the next round is a
 * guess, which is what the replan failure category describes ("第一轮失败后没有纠正").
 * The verifier's failure detail carries the cell, the expectation and the actual
 * value; this pins all three into the note the next round receives.
 */
test('a failing round tells the next round which assertion failed and why', async () => {
  const path = await makeWorkbook()
  const seen: AgentPlanContext[] = []
  let calls = 0
  const planner: AgentPlanner = {
    async plan(context) {
      calls++
      seen.push(context)
      // Round 1 asserts a value the workbook does not hold.
      if (calls === 1) {
        return {
          steps: [{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }],
          assertions: [{ id: '订单!B3', expect: 999 }],
        }
      }
      return {
        steps: [
          { name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] },
          { name: 'assert', operations: [{ op: 'set', cells: { '订单!E1': '已补齐' } }] },
        ],
        assertions: [{ id: '订单!B3', expect: 0 }],
      }
    },
    async verify() {
      return { achieved: true, reason: 'LLM 认为完成' }
    },
  }

  const result = await runAgentTask(path, { goal: '补空值', planner, maxRounds: 2 })
  assert.ok(seen[1], 'the loop should have run a second round')
  const feedback = seen[1]!.verifierNote ?? ''

  assert.match(feedback, /订单!B3/, 'the next round must be told which cell failed')
  assert.match(feedback, /999/, 'and what was expected')
  assert.match(feedback, /0/, 'and what was actually there')
  assert.equal(result.achieved, true)
})

test('a round that changes nothing does not undo an earlier round that did the work', async () => {
  // The 11 tasks this reproduces: the workbook ends up correct, every assertion passes,
  // and the run still reports achieved=false. The cause is that `changed` compares the
  // round's output against that round's own input, so a final round whose plan is a no-op
  // reads as "nothing happened" — even when the previous round already finished the job.
  //
  // The guard itself is worth keeping: a run where nothing ever changed has not achieved
  // anything, and calling that success would be worse than the bug. What is wrong is
  // scoping it to one round when the question is about the whole run.
  const path = await makeWorkbook()
  const planner = fakePlanner(
    [
      // Round 1 does the work.
      [{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }],
      // Round 2 asks for the same thing again — a no-op, because it is already done.
      [{ name: 'fill again', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }],
    ],
    [
      { achieved: false, reason: '还没检查' },
      { achieved: true, reason: '完成' },
    ],
  )
  const result = await runAgentTask(path, { goal: '把金额空值填 0', planner, maxRounds: 2 })

  // The evidence says the goal was reached…
  assert.equal(result.finalAnomalies, 0)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(result.outputPath)
  assert.equal(workbook.getWorksheet('订单')!.getCell('B3').value, 0)

  // …so the run should say so. A final no-op round is not evidence against it.
  assert.equal(result.achieved, true)
  assert.equal(result.rounds.length, 2)
})

test('a run in which nothing ever changed is still not achieved', async () => {
  // The other half of the guard, and the reason it was not simply deleted. A no-op final
  // round must not sink a run that did the work — but a run that never changed anything at
  // all has not achieved the goal, and reporting success there would be a worse failure
  // than the one being fixed. `changed` now asks the run-level question; this pins that the
  // question still has a "no" answer.
  const path = await makeWorkbook()
  const planner = fakePlanner(
    [
      [{ name: 'noop', operations: [{ op: 'set', cells: { '订单!A1': '区域' } }] }],
      [{ name: 'noop again', operations: [{ op: 'set', cells: { '订单!A1': '区域' } }] }],
    ],
    [
      { achieved: true, reason: '完成' },
      { achieved: true, reason: '完成' },
    ],
  )
  const result = await runAgentTask(path, { goal: '什么都不做', planner, maxRounds: 2 })

  assert.equal(result.achieved, false, 'nothing changed, so nothing was achieved')
  assert.match(result.rounds.at(-1)!.verdict.reason, /没有实质变化/)
})

/**
 * The note for a rejected plan carries the plan itself.
 *
 * `计划无效：…缺少必填字段 from` says what was wrong but not what the model wrote, and a
 * rejected plan clears `previousPlan` — so without this the model had nothing to correct
 * against and repeated the same plan. The note is the only channel it has.
 */
test('the note for a rejected plan carries the plan itself', async () => {
  const path = await makeWorkbook()
  const planner = fakePlanner(
    [
      [{ name: 'bad', operations: [{ op: 'groupColumns', sheet: '订单' } as never] }],
      [{ name: 'fill', operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }] }],
    ],
    [{ achieved: true, reason: '完成' }],
  )
  const result = await runAgentTask(path, { goal: '按列分组', planner, maxRounds: 2 })
  assert.equal(result.achieved, true)
  const note = planner.seen[1]!.verifierNote ?? ''
  assert.match(note, /groupColumns 缺少必填字段 from/)
  assert.match(note, /"op":"groupColumns"/)
})
