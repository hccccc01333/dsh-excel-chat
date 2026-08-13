import { before, test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { readChartInfos } from '../src/charts.ts'
import { validateCharts } from '../src/chart-validator.ts'
import { readWorkbookCells } from '../src/workbook.ts'
import { buildChartWorkbook } from './chart-fixture.ts'

const fixturePath = fileURLToPath(new URL('../fixtures/charts.xlsx', import.meta.url))

before(async () => {
  await mkdir(fileURLToPath(new URL('../fixtures', import.meta.url)), { recursive: true })
  await writeFile(fixturePath, await buildChartWorkbook())
})

test('readChartInfos finds charts and their types/series', async () => {
  const charts = await readChartInfos(fixturePath)
  assert.equal(charts.length, 3)
  assert.equal(charts[0]!.type, 'lineChart')
  assert.equal(charts[0]!.series.length, 1)
  assert.equal(charts[0]!.series[0]!.values, 'Sales!$B$2:$B$4')
})

test('validateCharts flags structural anomalies but accepts a good chart', async () => {
  const cells = await readWorkbookCells(await readFile(fixturePath))
  const reports = validateCharts(await readChartInfos(fixturePath), cells)
  assert.equal(reports[0]!.anomalies.length, 0)
  const badKinds = reports[1]!.anomalies.map((a) => a.kind)
  assert.ok(badKinds.includes('multi-dimensional-range'))
  assert.ok(badKinds.includes('missing-cells'))
  assert.ok(badKinds.includes('missing-categories'))
  assert.ok(reports[2]!.anomalies.some((a) => a.kind === 'unsorted-dates'))
})

test('chart workbook fixture is still readable by ExcelJS', async () => {
  const cells = await readWorkbookCells(await readFile(fixturePath))
  assert.equal(cells['Sales!B2'], '100')
  assert.ok(cells['Sales!A2'].startsWith('2026-01-01'))
})
