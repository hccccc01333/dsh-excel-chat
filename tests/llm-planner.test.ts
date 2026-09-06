import { test } from 'node:test'
import assert from 'node:assert/strict'
import { createLlmPlanner } from '../src/llm-planner.ts'

test('verifier prompt requires evidence-based checklist', async () => {
  let captured = ''
  const planner = createLlmPlanner(async (prompt) => {
    captured = prompt
    return JSON.stringify({ achieved: true, reason: 'ok' })
  })
  await planner.verify({
    goal: '按区域汇总金额',
    path: 'D:/sales.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '订单：3 行 × 4 列',
    validationSummary: '0 个公式异常',
    previousPlan: undefined,
    previousResult: undefined,
    verifierNote: undefined,
    executedPlan: [{ name: '汇总', operations: [] }],
    executedResult: { outputPath: 'D:/out.xlsx', steps: [], finalAnomalies: 0 },
    cellSnapshot: '订单!A1=区域',
  })
  assert.match(captured, /可检查点/)
  assert.match(captured, /每一条都要有明确证据/)
})

test('verifier prompt declares snapshot sampling and per-goal evidence standards', async () => {
  let captured = ''
  const planner = createLlmPlanner(async (prompt) => {
    captured = prompt
    return JSON.stringify({ achieved: false, reason: 'no' })
  })
  await planner.verify({
    goal: '做交叉表',
    path: 'D:/sales.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '订单：3 行 × 3 列',
    validationSummary: '0 个公式异常',
    previousPlan: undefined,
    previousResult: undefined,
    verifierNote: undefined,
    executedPlan: [{ name: 'x', operations: [] }],
    executedResult: { outputPath: 'D:/out.xlsx', steps: [], finalAnomalies: 0 },
    cellSnapshot: '订单!A1=区域',
  })
  assert.match(captured, /按工作表轮询采样/)
  assert.match(captured, /SUMIFS/)
  assert.match(captured, /反例/)
})

test('planner prompt includes new ops, few-shot examples, and analysis rules', async () => {
  let captured = ''
  const planner = createLlmPlanner(async (prompt) => {
    captured = prompt
    return JSON.stringify({ steps: [{ name: 's', operations: [{ op: 'aggregateReport' }] }] })
  })
  const plan = await planner.plan({
    goal: '按区域汇总金额',
    path: 'D:/sales.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '订单：3 行 × 3 列',
    validationSummary: '0 个公式异常',
  })
  const steps = Array.isArray(plan) ? plan : plan.steps
  assert.match(captured, /crosstab/)
  assert.match(captured, /joinSheets/)
  assert.match(captured, /rankColumn/)
  assert.match(captured, /完整示例/)
  assert.match(captured, /禁止用 set 写死汇总数字/)
  assert.match(captured, /metric 是对象/)
  assert.match(captured, /assertions/)
  assert.match(captured, /机器断言/)
})

test('planner normalizes crosstab flat metric and single-string arrays', async () => {
  const planner = createLlmPlanner(async () => JSON.stringify({
    steps: [{
      name: 's',
      operations: [
        { op: 'crosstab', source: '订单!A1:C4', rowColumn: 'A', columnColumn: 'B', metricColumn: 'C', metricFunction: 'average' },
        { op: 'joinSheets', source: '订单!A1:B4', sourceKey: 'A', lookup: '客户!A1:C3', lookupKey: 'A', valueColumns: 'B', outputColumns: 'C' },
      ],
    }],
  }))
  const plan = await planner.plan({
    goal: '交叉表',
    path: 'D:/x.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '',
    validationSummary: '',
  })
  const steps = Array.isArray(plan) ? plan : plan.steps
  const crosstab = steps[0]!.operations[0] as Record<string, unknown>
  assert.deepEqual(crosstab.metric, { column: 'C', function: 'average' })
  const join = steps[0]!.operations[1] as Record<string, unknown>
  assert.deepEqual(join.valueColumns, ['B'])
  assert.deepEqual(join.outputColumns, ['C'])
})

test('planner passes machine assertions through when the model emits them', async () => {
  const planner = createLlmPlanner(async () => JSON.stringify({
    steps: [{ name: 's', operations: [{ op: 'aggregateReport', source: '订单!A1:C4', groupColumn: 'A', metrics: [{ column: 'C', function: 'sum' }] }] }],
    assertions: [{ id: '汇总!B2', startsWith: '=SUMIFS(' }],
  }))
  const plan = await planner.plan({
    goal: '按区域汇总金额',
    path: 'D:/x.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '',
    validationSummary: '',
  })
  assert.ok(!Array.isArray(plan))
  assert.deepEqual(plan.assertions, [{ id: '汇总!B2', startsWith: '=SUMIFS(' }])
})
