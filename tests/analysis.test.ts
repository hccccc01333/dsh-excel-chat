import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildCorpusWorkbook } from '../src/corpus/helpers.ts'
import { analyzeWorkbook, type AnalysisFinding } from '../src/analysis.ts'

/**
 * The analytical findings.
 *
 * Every expectation here came from running the thing, not from reading it. The trend
 * logic went through three corrections, each found by looking at the output rather than
 * at a pass/fail:
 *
 *   1. it never fired for three-period series — the header row was counted as a
 *      non-period value, so 3 of 4 values were periods and the 80% bar failed
 *   2. it reported "upward" with a negative change — direction came from counting steps,
 *      the change from comparing the ends, and a late collapse made them disagree
 *   3. it reported "upward +50%" for 100 → 300 → 600 → 900 → 150 — arithmetically true,
 *      and misleading, because the series collapsed from its peak
 *
 * The invariant that caught (2) and (3) is asserted on every finding below.
 */
const sheet = (rows: Array<[string, number]>, headers = ['月份', '金额']) => [{ name: '销售', headers, rows }]

async function findings(rows: Array<[string, number]>, dir: string): Promise<AnalysisFinding[]> {
  const path = await buildCorpusWorkbook(dir, `t${Math.random().toString(36).slice(2, 8)}`, sheet(rows))
  return analyzeWorkbook(path)
}

const trendOf = (all: AnalysisFinding[]) => all.find((finding) => finding.kind === 'trend')

/**
 * A trend finding may never contradict itself.
 *
 * Stated once and applied to every case: the direction it names has to match the sign of
 * the change it prints. Both bugs above would have failed here.
 */
function assertSelfConsistent(finding: AnalysisFinding): void {
  const direction = finding.evidence.direction
  const change = Number(finding.evidence.changePercent)
  assert.equal(
    direction,
    change >= 0 ? 'up' : 'down',
    `says "${direction}" but reports ${change}%: ${finding.message}`,
  )
  const last = Number(finding.evidence.lastTotal)
  const first = Number(finding.evidence.firstTotal)
  assert.equal(
    direction,
    last >= first ? 'up' : 'down',
    `direction disagrees with the endpoints it prints: ${finding.message}`,
  )
}

test('reports a trend across a short series', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  const trend = trendOf(await findings([['2026-01', 100], ['2026-02', 200], ['2026-03', 350]], dir))
  assert.ok(trend, 'a three-period ramp is a trend')
  assertSelfConsistent(trend)
  assert.equal(trend.evidence.direction, 'up')
  assert.equal(trend.evidence.firstTotal, 100)
  assert.equal(trend.evidence.lastTotal, 350)
})

test('a single dip does not disqualify a trend', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  const trend = trendOf(await findings(
    [['2026-01', 100], ['2026-02', 130], ['2026-03', 120], ['2026-04', 210], ['2026-05', 260]],
    dir,
  ))
  assert.ok(trend, 'rising with one bad month is still rising')
  assertSelfConsistent(trend)
  assert.equal(trend.evidence.consistentSteps, 3)
  assert.equal(trend.evidence.steps, 4)
})

test('a series that collapses from its peak is not an upward trend', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  // Net +50%, and three of four steps up — arithmetically an upward trend, and the last
  // thing a reader would conclude from the shape.
  const trend = trendOf(await findings(
    [['2026-01', 100], ['2026-02', 300], ['2026-03', 600], ['2026-04', 900], ['2026-05', 150]],
    dir,
  ))
  assert.equal(trend, undefined)
})

test('a series that wanders is not a trend', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  const trend = trendOf(await findings(
    [['2026-01', 100], ['2026-02', 300], ['2026-03', 80], ['2026-04', 250], ['2026-05', 90]],
    dir,
  ))
  assert.equal(trend, undefined)
})

test('two periods are not enough to call a trend', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  const trend = trendOf(await findings([['2026-01', 100], ['2026-02', 300]], dir))
  assert.equal(trend, undefined)
})

test('a downward series reports as downward', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  const trend = trendOf(await findings(
    [['2026-01', 500], ['2026-02', 420], ['2026-03', 330], ['2026-04', 200]],
    dir,
  ))
  assert.ok(trend)
  assertSelfConsistent(trend)
  assert.equal(trend.evidence.direction, 'down')
})

test('a late rebound is not a downward trend', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  const trend = trendOf(await findings(
    [['2026-01', 500], ['2026-02', 420], ['2026-03', 330], ['2026-04', 380]],
    dir,
  ))
  assert.equal(trend, undefined)
})

test('quarter and Chinese month labels sort as periods, not as text', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  // `10月` sorts before `3月` as a string; the series has to be ordered by period or the
  // trend is computed over a meaningless order.
  const chinese = trendOf(await findings([['1月', 100], ['2月', 130], ['3月', 165], ['10月', 260]], dir))
  assert.ok(chinese)
  assertSelfConsistent(chinese)
  assert.equal(chinese.evidence.firstPeriod, '1月')
  assert.equal(chinese.evidence.lastPeriod, '10月')

  const quarterly = trendOf(await findings([['2026-Q1', 100], ['2026-Q2', 200], ['2026-Q3', 350]], dir))
  assert.ok(quarterly)
  assertSelfConsistent(quarterly)
})

test('reports a measure concentrated in a few categories', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  const path = await buildCorpusWorkbook(dir, 'conc', [{
    name: '销售',
    headers: ['区域', '金额'],
    rows: [['华北', 900], ['华南', 870], ['华东', 120], ['西南', 90]],
  }])
  const concentration = (await analyzeWorkbook(path)).find((finding) => finding.kind === 'concentration')
  assert.ok(concentration, 'one region at 45% of the total is worth saying')
  assert.equal(concentration.evidence.top, '华北')
  assert.equal(concentration.evidence.categories, 4)
  // The share it prints has to match the numbers it prints.
  const total = Number(concentration.evidence.total)
  assert.ok(Math.abs(Number(concentration.evidence.topShare) - (900 / total) * 100) < 0.1)
})

test('a balanced table is not reported as concentrated', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  const path = await buildCorpusWorkbook(dir, 'balanced', [{
    name: '销售',
    headers: ['区域', '金额'],
    rows: [['华北', 250], ['华南', 250], ['华东', 250], ['西南', 250]],
  }])
  const concentration = (await analyzeWorkbook(path)).find((finding) => finding.kind === 'concentration')
  assert.equal(concentration, undefined, 'four equal categories is what balanced looks like')
})

test('a sheet with no time or dimension column yields nothing rather than a guess', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'analysis-'))
  const path = await buildCorpusWorkbook(dir, 'flat', [{
    name: '备注',
    headers: ['说明'],
    rows: [['a'], ['b'], ['c'], ['d']],
  }])
  assert.deepEqual(await analyzeWorkbook(path), [])
})
