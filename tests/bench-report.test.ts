import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFileSync, readdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import { build, validateRun } from '../scripts/bench-report.mjs'

const ROOT = fileURLToPath(new URL('..', import.meta.url))

/**
 * The published table is generated from the run artifacts, so the failure mode to
 * guard is a table that no longer matches its data — the same drift the images had.
 */
test('the committed results table matches the run artifacts', () => {
  const { markdown } = build()
  const committed = readFileSync(join(ROOT, 'docs', 'benchmark-results.md'), 'utf8')
  assert.equal(committed, markdown, 'docs/benchmark-results.md is stale; run `npm run bench:report`')
})

test('every run artifact is comparable: it names a model and a date', () => {
  const dir = join(ROOT, 'bench-results')
  const files = readdirSync(dir).filter((name) => name.endsWith('.json'))
  assert.ok(files.length > 0, 'no run artifacts at all')
  for (const file of files) {
    const run = validateRun(file, JSON.parse(readFileSync(join(dir, file), 'utf8')))
    assert.ok(run.model, `${file} has no model`)
    assert.ok(run.date, `${file} has no date`)
    assert.ok(run.rate > 0 && run.rate <= 1, `${file} has a nonsensical success rate`)
    // A transcribed row must say so, or it reads as machine-generated.
    assert.ok(run.provenance, `${file} has no provenance note`)
  }
})

test('a run that never reached the model is rejected, not published as 0%', () => {
  // This is exactly what `node tests/invoke-llm-benchmark.ts` produces without an
  // API key: every task recorded as an execution failure.
  const keyless = {
    model: 'deepseek-chat',
    date: '2026-10-08',
    total: 100,
    success: 0,
    successRate: 0,
    failureBreakdown: { execution: 100 },
  }
  assert.throws(() => validateRun('keyless.json', keyless), /no API key/)

  // A genuine zero from a model that ran is still publishable.
  const realZero = { ...keyless, failureBreakdown: { execution: 40, verification: 60 } }
  assert.equal(validateRun('real-zero.json', realZero).rate, 0)
})

test('a run without a model or date is rejected rather than shown as a blank cell', () => {
  assert.throws(() => validateRun('x.json', { total: 10, success: 5, date: '2026-10-08' }), /no model/)
  assert.throws(() => validateRun('x.json', { total: 10, success: 5, model: 'm' }), /no date/)
  assert.throws(() => validateRun('x.json', { model: 'm', date: '2026-10-08' }), /no successRate/)
})
