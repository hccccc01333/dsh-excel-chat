import { test } from 'node:test'
import assert from 'node:assert/strict'
import { classifyFailure, summarizeFailureBreakdown, type FailureEvidence } from '../src/failure-taxonomy.ts'
import { runLlmTask } from '../src/llm-benchmark.ts'
import type { AgentPlanner } from '../src/agent.ts'
import type { FileBenchmarkTask } from '../src/file-benchmark.ts'
import { makeSalesWorkbook } from './sales-fixture.ts'

function evidence(overrides: Partial<FailureEvidence>): FailureEvidence {
  return {
    crashed: false,
    error: null,
    verifierFalsePositive: false,
    rounds: 1,
    maxRounds: 1,
    executedOps: [],
    expectedOps: [],
    argDiffs: [],
    checksPassed: 0,
    checksTotal: 2,
    integrity: 1,
    ...overrides,
  }
}

test('classifyFailure maps a crashed engine error to execution', () => {
  const result = classifyFailure(evidence({ crashed: true, error: 'ENOENT: no such file' }))
  assert.equal(result.category, 'execution')
})

test('classifyFailure maps an invalid plan to planning', () => {
  const result = classifyFailure(evidence({ crashed: true, error: 'planner returned an empty plan' }))
  assert.equal(result.category, 'planning')
})

test('classifyFailure maps a sanitize-rejected plan to planning', () => {
  const result = classifyFailure(evidence({ crashed: true, error: 'conditionalFormatting 缺少必填数组 rules' }))
  assert.equal(result.category, 'planning')
})

test('classifyFailure maps an invalid range to argument', () => {
  const result = classifyFailure(evidence({
    crashed: true,
    error: 'invalid range: 区域汇总!A3（第 3 轮计划：fillSeries(target="区域汇总!A3")）',
  }))
  assert.equal(result.category, 'argument')
})

test('classifyFailure maps a verifier false positive to verification', () => {
  const result = classifyFailure(evidence({ verifierFalsePositive: true, checksPassed: 1 }))
  assert.equal(result.category, 'verification')
})

test('classifyFailure maps exhausted replanning to replan', () => {
  const result = classifyFailure(evidence({ rounds: 3, maxRounds: 3 }))
  assert.equal(result.category, 'replan')
})

test('classifyFailure maps an empty plan to planning', () => {
  const result = classifyFailure(evidence({ rounds: 1, maxRounds: 1, executedOps: [] }))
  assert.equal(result.category, 'planning')
})

test('classifyFailure maps wrong arguments to argument', () => {
  const result = classifyFailure(evidence({
    executedOps: ['aggregateReport'],
    expectedOps: ['aggregateReport'],
    argDiffs: ['aggregateReport.metrics：期望 [{"column":"F","function":"sum"}]，实际 [{"column":"E","function":"sum"}]'],
  }))
  assert.equal(result.category, 'argument')
})

test('classifyFailure maps matching ops with failing checks to semantic', () => {
  const result = classifyFailure(evidence({
    executedOps: ['aggregateReport'],
    expectedOps: ['aggregateReport'],
    argDiffs: [],
  }))
  assert.equal(result.category, 'semantic')
})

test('classifyFailure maps a generic fallback to tool-selection', () => {
  const result = classifyFailure(evidence({
    executedOps: ['set'],
    expectedOps: ['aggregateReport'],
  }))
  assert.equal(result.category, 'tool-selection')
})

test('classifyFailure maps a missing step to planning', () => {
  const result = classifyFailure(evidence({
    executedOps: ['style'],
    expectedOps: ['sortRange', 'style'],
  }))
  assert.equal(result.category, 'planning')
})

test('classifyFailure maps unrelated operations to intent', () => {
  const result = classifyFailure(evidence({
    executedOps: ['deleteSheet'],
    expectedOps: ['report'],
  }))
  assert.equal(result.category, 'intent')
})

test('summarizeFailureBreakdown counts categories', () => {
  const breakdown = summarizeFailureBreakdown([
    { category: 'verification' },
    { category: 'argument' },
    { category: 'verification' },
  ])
  assert.equal(breakdown.verification, 2)
  assert.equal(breakdown.argument, 1)
  assert.equal(breakdown.intent, 0)
})

function plannerReturning(
  plan: Array<{ operations: Array<Record<string, unknown>> }>,
  verdict: { achieved: boolean; reason: string } = { achieved: true, reason: '目标已达成' },
): AgentPlanner {
  return {
    async plan() {
      return plan as never
    },
    async verify() {
      return verdict
    },
  }
}

function cellTask(expectValue: string): FileBenchmarkTask {
  return {
    id: 'taxonomy-wiring',
    category: 'editing',
    name: 'taxonomy wiring',
    description: '把 订单!A1 改成指定值',
    buildInput: () => makeSalesWorkbook(),
    operations: [{ op: 'set', cells: { '订单!A1': expectValue } }],
    checks: [{ id: '订单!A1', expect: expectValue }],
  }
}

test('runLlmTask classifies a verifier false positive end-to-end', async () => {
  const task = cellTask('x')
  const planner = plannerReturning([{ operations: [{ op: 'set', cells: { '订单!A1': 'WRONG' } }] }])
  const result = await runLlmTask(task, { planner, maxRounds: 2 })
  assert.equal(result.success, false)
  assert.equal(result.achieved, true)
  assert.equal(result.failure?.category, 'verification')
})

test('runLlmTask classifies an argument mismatch end-to-end', async () => {
  const task = cellTask('x')
  const planner = plannerReturning(
    [{ operations: [{ op: 'set', cells: { '订单!A1': 'y' } }] }],
    { achieved: false, reason: '未完成' },
  )
  const result = await runLlmTask(task, { planner, maxRounds: 1 })
  assert.equal(result.success, false)
  assert.equal(result.failure?.category, 'argument')
})
