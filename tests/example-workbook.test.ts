import { test } from 'node:test'
import assert from 'node:assert/strict'
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { autofixWorkbookFile } from '../src/autofix.ts'
import { readWorkbookCells, validateWorkbookFile } from '../src/workbook.ts'
import { EXAMPLE_PATH, EXPECTED_FORMULA, writeExample } from '../scripts/make-example.mjs'
import { runTour, renderTour } from '../scripts/demo.mjs'

/**
 * The sample workbook is a front door, and a front door that has quietly turned into a
 * clean file is worse than no sample at all: the reader follows the walkthrough, watches
 * nothing happen, and concludes the tool does nothing. So the file's entire reason for
 * existing is asserted here — one detectable anomaly, one deterministic repair, a clean
 * re-check — rather than trusted to stay true.
 */
test('the shipped example still demonstrates the loop it exists for', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-example-'))
  try {
    const working = join(dir, 'working.xlsx')
    await copyFile(EXAMPLE_PATH, working)

    const validation = await validateWorkbookFile(working)
    assert.equal(validation.anomalies.length, 1, 'the example should carry exactly one anomaly')
    const anomaly = validation.anomalies[0]!
    assert.equal(anomaly.kind, 'reference-offset')
    assert.equal(anomaly.cell, '销售台账!G9')

    const fix = await autofixWorkbookFile(working, { outPath: join(dir, 'repaired.xlsx'), healthReport: false })
    assert.deepEqual(
      fix.repairs.map((patch) => `${patch.oldValue} -> ${patch.newValue}`),
      ['=E9*F8 -> =E9*F9'],
    )
    assert.equal(fix.repairs[0]!.newValue, EXPECTED_FORMULA)
    // Read the repaired file back. The repair step's own summary is a claim, not evidence.
    assert.equal((await validateWorkbookFile(fix.repairedPath)).anomalies.length, 0)
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
})

/**
 * And the committed file has to be the one the generator produces.
 *
 * Compared as cell content rather than bytes: a zip carries timestamps, so two identical
 * workbooks never hash the same. A hand-edit to the committed workbook — the tempting way
 * to "fix" the sample — shows up here instead of silently changing what the walkthrough
 * demonstrates.
 */
test('the committed example matches what the generator produces', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-example-src-'))
  try {
    const regenerated = await writeExample(join(dir, 'sales-ledger.xlsx'))
    assert.deepEqual(
      await readWorkbookCells(await readFile(regenerated)),
      await readWorkbookCells(await readFile(EXAMPLE_PATH)),
    )
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
})

/**
 * The README's walkthrough is `npm run demo`, so what has to stay true is the *printed*
 * text: five stages, in order, each carrying the engine's own output.
 *
 * Asserting on the rendered string rather than on the steps behind it is deliberate. A
 * refactor that kept the data but stopped printing the repair — or that quietly dropped
 * the edit and left four stages — would otherwise pass, and the first person to notice
 * would be a reader who ran it and saw the promise not kept.
 */
test('the walkthrough prints the whole loop, and the loop really closes', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-demo-'))
  try {
    const tour = await runTour(dir)
    const text = renderTour(tour)

    // Every stage is present, in order, and the first is a real edit rather than a greeting.
    let cursor = -1
    for (const stage of ['① 改表', '② 体检', '③ 修复', '④ 复验', '⑤ 输出']) {
      const at = text.indexOf(stage)
      assert.ok(at > cursor, `missing or out-of-order stage: ${stage}`)
      cursor = at
    }

    // The edit produced the summary sheet the README promises, groups and all.
    assert.deepEqual(tour.groups, ['华东', '华北', '华南', '总计'])
    assert.match(text, /销售台账-汇总：A1「区域」 B1「金额 合计」/)

    // Detection found the sample's one planted error, and named the cell.
    assert.equal(tour.edit.validation.anomalies.length, 1)
    assert.equal(tour.edit.validation.anomalies[0]!.cell, '销售台账!G9')
    assert.match(text, /销售台账!G9 {2}\[reference-offset\]/)

    // The repair is the real one, and the re-check ran on the repaired file.
    assert.deepEqual(
      tour.repairs.map((patch) => `${patch.oldValue} -> ${patch.newValue}`),
      ['=E9*F8 -> =E9*F9'],
    )
    assert.equal(tour.final.anomalies.length, 0)
    assert.match(text, /异常 0 —— 公式列重新一致/)
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
})
