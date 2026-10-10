import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AgentPlanner } from '../src/agent.ts'
import { corpusTasks } from '../src/corpus/index.ts'
import { buildCorpusWorkbook } from '../src/corpus/helpers.ts'
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

test('the benchmark runs tasks concurrently and still reports them in order', async () => {
  // Serial, a full corpus took hours while the model answered in about four seconds —
  // nearly all of it spent waiting. Tasks are independent, so concurrency changes the
  // clock and nothing else; this test pins both halves of that claim.
  const seen: string[] = []
  const inFlight = { current: 0, peak: 0 }
  const tasks = Array.from({ length: 6 }, (_, index) => ({
    id: `t${index}`,
    category: 'editing',
    name: `task ${index}`,
    description: '',
    // A real workbook: the runner loads the input before it ever asks the planner, so a
    // stub path would fail first and the overlap this test is about would never happen.
    buildInput: (dir: string) =>
      buildCorpusWorkbook(dir, `t${index}`, [
        { name: 'Sheet1', headers: ['a', 'b'], rows: [['x', index]] },
      ]),
    operations: [],
    checks: [],
  })) as unknown as FileBenchmarkTask[]

  const options = {
    planner: {
      plan: async () => {
        inFlight.current += 1
        inFlight.peak = Math.max(inFlight.peak, inFlight.current)
        // Long enough that a serial loop could never overlap two of them.
        await new Promise((resolve) => setTimeout(resolve, 60))
        inFlight.current -= 1
        seen.push('call')
        return [{ name: 'noop', operations: [] }]
      },
      verify: async () => ({ achieved: true, reason: 'ok' }),
    },
    maxRounds: 1,
  } as unknown as LlmBenchmarkOptions

  const report = await runLlmBenchmark(tasks, { ...options, concurrency: 3 })

  assert.ok(inFlight.peak > 1, `tasks did not overlap; peak in flight was ${inFlight.peak}`)
  assert.ok(inFlight.peak <= 3, `concurrency limit exceeded: ${inFlight.peak} in flight`)
  // Order is the caller's, not the completion order — a shuffled report would make two
  // runs of the same corpus diff badly.
  assert.deepEqual(
    report.tasks.map((task) => task.id),
    tasks.map((task) => task.id),
  )
})
