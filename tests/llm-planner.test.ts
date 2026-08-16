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
