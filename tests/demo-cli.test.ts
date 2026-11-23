import assert from 'node:assert/strict'
import { test } from 'node:test'
import { execFile } from 'node:child_process'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

const ROOT = fileURLToPath(new URL('..', import.meta.url))
const DEMO = fileURLToPath(new URL('../scripts/demo.mjs', import.meta.url))
const SAMPLE = fileURLToPath(new URL('../examples/sales-ledger.xlsx', import.meta.url))

/**
 * The other demo tests import `runTour`/`renderTour`, so a broken CLI entry
 * point (or a stray print on import) cannot fail the suite. Run the real
 * command: it is the one a new user copies from the README.
 */
test('npm run demo prints all five stages and leaves the example untouched', async () => {
  const before = createHash('sha256').update(await readFile(SAMPLE)).digest('hex')
  const stdout = await new Promise<string>((resolve, reject) => {
    execFile(process.execPath, [DEMO], { cwd: ROOT, encoding: 'utf8' }, (error, out, stderr) => {
      if (error !== null) reject(new Error(`demo failed: ${error.message}\n${stderr}`))
      else resolve(out)
    })
  })
  for (const stage of ['① 改表', '② 体检', '③ 修复', '④ 复验', '⑤ 输出']) {
    assert.ok(stdout.includes(stage), `the demo output is missing ${stage}`)
  }
  assert.match(stdout, /未自动修复|异常 0/, 'the demo must say what happened to unrepaired anomalies')
  const after = createHash('sha256').update(await readFile(SAMPLE)).digest('hex')
  assert.equal(after, before, 'the demo rewrote the checked-in example workbook')
})
