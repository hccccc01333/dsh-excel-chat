import { test } from 'node:test'
import assert from 'node:assert/strict'
import type { AgentPlanner } from '../src/agent.ts'
import { corpusTasks } from '../src/corpus/index.ts'
import { runLlmBenchmark, runLlmTask } from '../src/llm-benchmark.ts'

test('LLM benchmark runner scores the canonical plan 100% offline', async () => {
  const plannerFor = (task: (typeof corpusTasks)[number]): AgentPlanner => ({
    plan: async () => [{ name: 'canonical', operations: task.operations }],
    verify: async () => ({ achieved: true, reason: '标准答案' }),
  })
  const results = []
  for (const task of corpusTasks) {
    results.push(await runLlmTask(task, { planner: plannerFor(task), maxRounds: 2 }))
  }
  const failed = results.filter((result) => !result.success)
  assert.deepEqual(failed, [])
})

test('LLM benchmark report aggregates categories', async () => {
  const tasks = corpusTasks.slice(0, 4)
  const planner: AgentPlanner = {
    plan: async () => [{ name: 'x', operations: [] }],
    verify: async () => ({ achieved: true, reason: 'x' }),
  }
  const report = await runLlmBenchmark(tasks, { planner })
  assert.equal(report.total, 4)
  assert.ok(report.categories.editing.total >= 1)
})
