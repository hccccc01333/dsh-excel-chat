import assert from 'node:assert/strict'
import { test } from 'node:test'
import { existsSync } from 'node:fs'
import { mkdtemp, readdir, readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { buildCorpusWorkbook } from '../src/corpus/helpers.ts'
import { runExcelTask } from '../src/task.ts'

/**
 * Scratch directories are removed, and the deliverable is not.
 *
 * `runExcelTask` and `runAgentTask` each create a temporary directory for their
 * intermediate workbooks and neither removed it. One leak per call is invisible in a
 * test and adds up in a long-lived session, which is where this plugin runs — on the
 * machine this was written on there were **21,655** leftover `vera-task-*` directories
 * waiting in `%TEMP%`.
 *
 * The second half matters as much as the first: the finished file is copied out of the
 * scratch directory before the cleanup runs, and a cleanup that took the output with it
 * would be far worse than the leak.
 *
 * The temp root is redirected to a directory of this file's own. `node --test` runs test
 * files concurrently, so any other file that calls into `task.ts` drops scratch
 * directories into the shared temp root while this one is counting — which is exactly
 * how the first version of this test failed. `os.tmpdir()` re-reads the environment on
 * every call, and each test file has its own process, so this is private in practice.
 */
const PRIVATE_TMP = await mkdtemp(join(process.env.TEMP ?? process.env.TMPDIR ?? '.', 'scratch-cleanup-root-'))
process.env.TMPDIR = PRIVATE_TMP
process.env.TMP = PRIVATE_TMP
process.env.TEMP = PRIVATE_TMP

/** Scratch directories this test's own temp root holds, by name. */
async function scratchDirs(prefix: string): Promise<Set<string>> {
  return new Set((await readdir(PRIVATE_TMP)).filter((name) => name.startsWith(prefix)))
}

async function fixture(): Promise<string> {
  const dir = await mkdtemp(join(PRIVATE_TMP, 'scratch-source-'))
  return buildCorpusWorkbook(dir, 'scratch', [
    { name: '订单', headers: ['产品', '数量'], rows: [['苹果', 10], ['香蕉', 5]] },
  ])
}

test('a completed task removes its scratch directory and keeps its output', async () => {
  const before = await scratchDirs('vera-task-')
  const path = await fixture()
  const out = join(PRIVATE_TMP, `task-out-${Date.now()}.xlsx`)

  const result = await runExcelTask(
    path,
    [{ name: '改一个值', operations: [{ op: 'set', cells: { '订单!B2': '11' } }] }],
    out,
  )

  // The deliverable is there…
  assert.ok(existsSync(result.outputPath), 'the task output was not written')
  assert.equal(result.outputPath, out)
  assert.ok((await readFile(result.outputPath)).length > 0, 'the output file is empty')

  // …and nothing was left in the temp directory.
  const leaked = [...(await scratchDirs('vera-task-'))].filter((name) => !before.has(name))
  assert.deepEqual(leaked, [], 'the scratch directory survived a successful task')
})

test('a failed task removes its scratch directory too', async () => {
  const before = await scratchDirs('vera-task-')
  const path = await fixture()

  // A step that cannot succeed: the sheet does not exist. The point is that the cleanup
  // runs on the throwing path as well, which is the path a leak hides on.
  await assert.rejects(() =>
    runExcelTask(path, [{ operations: [{ op: 'set', cells: { '不存在的表!A1': 'x' } }] }]),
  )

  const leaked = [...(await scratchDirs('vera-task-'))].filter((name) => !before.has(name))
  assert.deepEqual(leaked, [], 'the scratch directory survived a failed task')
})

test('repeated runs do not accumulate scratch directories', async () => {
  // One run leaking is a detail; the leak this guards against is the one that scales —
  // and the one that put 21,655 directories in a real temp folder.
  const before = await scratchDirs('vera-task-')
  const path = await fixture()
  for (let index = 0; index < 3; index++) {
    await runExcelTask(
      path,
      [{ operations: [{ op: 'set', cells: { '订单!B2': String(index + 1) } }] }],
      join(PRIVATE_TMP, `run-${index}.xlsx`),
    )
  }
  const leaked = [...(await scratchDirs('vera-task-'))].filter((name) => !before.has(name))
  assert.deepEqual(leaked, [], 'three runs left three scratch directories behind')
})
