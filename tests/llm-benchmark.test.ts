import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AgentPlanner } from '../src/agent.ts'
import { corpusTasks } from '../src/corpus/index.ts'
import { readBenchmarkCheckpoint, runLlmBenchmark, runLlmTask } from '../src/llm-benchmark.ts'

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

test('runLlmBenchmark checkpoints completed tasks to a JSONL file', async () => {
  const tasks = corpusTasks.slice(0, 3)
  const dir = await mkdtemp(join(tmpdir(), 'vera-bench-checkpoint-'))
  const outFile = join(dir, 'run.jsonl')
  let calls = 0
  const planner: AgentPlanner = {
    plan: async () => {
      calls++
      return [{ name: 'x', operations: [] }]
    },
    verify: async () => ({ achieved: true, reason: 'x' }),
  }
  const first = await runLlmBenchmark(tasks, { planner, outFile })
  assert.equal(first.total, 3)
  const callsAfterFirst = calls
  assert.ok(callsAfterFirst >= 3)
  const lines = (await readFile(outFile, 'utf8')).trim().split('\n')
  assert.equal(lines.length, 3)
  const checkpoint = await readBenchmarkCheckpoint(outFile)
  assert.deepEqual([...checkpoint.keys()].sort(), tasks.map((task) => task.id).sort())

  // Rerun over the same file: all tasks come from the checkpoint, planner untouched.
  const second = await runLlmBenchmark(tasks, { planner, outFile })
  assert.equal(second.total, 3)
  assert.equal(second.successRate, first.successRate)
  assert.equal(calls, callsAfterFirst)
})

test('runLlmBenchmark resumes across slices and merges the report', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-bench-resume-'))
  const outFile = join(dir, 'slices.jsonl')
  const canonical: AgentPlanner = {
    plan: async () => [{ name: 'canonical', operations: [] }],
    verify: async () => ({ achieved: true, reason: 'x' }),
  }
  const sliceA = corpusTasks.slice(0, 2)
  const sliceB = corpusTasks.slice(2, 5)
  await runLlmBenchmark(sliceA, { planner: canonical, outFile })
  await runLlmBenchmark(sliceB, { planner: canonical, outFile })
  // Re-running the first slice through the same checkpoint still yields the
  // merged corpus results for those ids.
  const merged = await runLlmBenchmark(corpusTasks.slice(0, 5), { planner: canonical, outFile })
  assert.equal(merged.total, 5)
  assert.deepEqual(
    merged.tasks.map((task) => task.id).sort(),
    corpusTasks.slice(0, 5).map((task) => task.id).sort(),
  )
})
