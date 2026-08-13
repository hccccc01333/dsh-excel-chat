import { before, test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { strToU8, unzipSync, zipSync } from 'fflate'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { readChartInfos } from '../src/charts.ts'
import { validateCharts } from '../src/chart-validator.ts'
import { readWorkbookCells } from '../src/workbook.ts'

const fixturePath = fileURLToPath(new URL('../fixtures/charts.xlsx', import.meta.url))

before(async () => {
  await mkdir(fileURLToPath(new URL('../fixtures', import.meta.url)), { recursive: true })
  await writeFile(fixturePath, await buildChartWorkbook())
})

function chartXml(type: string, seriesXml: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:layout/><c:${type}>${seriesXml}</c:${type}><c:catAx/><c:valAx/></c:plotArea></c:chart>
</c:chartSpace>`
}

const goodSeries = `<c:ser>
  <c:tx><c:strRef><c:f>Sales!$B$1</c:f></c:strRef></c:tx>
  <c:cat><c:strRef><c:f>Sales!$A$2:$A$4</c:f></c:strRef></c:cat>
  <c:val><c:numRef><c:f>Sales!$B$2:$B$4</c:f></c:numRef></c:val>
</c:ser>`

const badSeries = `<c:ser>
  <c:cat><c:strRef><c:f>Sales!$A$2:$C$4</c:f></c:strRef></c:cat>
  <c:val><c:numRef><c:f>Sales!$X$9:$X$10</c:f></c:numRef></c:val>
</c:ser>
<c:ser>
  <c:val><c:numRef><c:f>Sales!$B$2:$B$4</c:f></c:numRef></c:val>
</c:ser>`

const unsortedSeries = `<c:ser>
  <c:cat><c:strRef><c:f>Sales!$D$2:$D$4</c:f></c:strRef></c:cat>
  <c:val><c:numRef><c:f>Sales!$B$2:$B$4</c:f></c:numRef></c:val>
</c:ser>`

async function buildChartWorkbook(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sales')
  sheet.getCell('A1').value = 'Date'
  sheet.getCell('A2').value = new Date('2026-01-01T00:00:00Z')
  sheet.getCell('A3').value = new Date('2026-02-01T00:00:00Z')
  sheet.getCell('A4').value = new Date('2026-03-01T00:00:00Z')
  sheet.getCell('D2').value = new Date('2026-03-01T00:00:00Z')
  sheet.getCell('D3').value = new Date('2026-02-01T00:00:00Z')
  sheet.getCell('D4').value = new Date('2026-01-01T00:00:00Z')
  sheet.getCell('B1').value = 'Revenue'
  sheet.getCell('B2').value = 100
  sheet.getCell('B3').value = 200
  sheet.getCell('B4').value = 300
  sheet.getCell('C2').value = 60
  sheet.getCell('C3').value = 120
  sheet.getCell('C4').value = 150
  const files = unzipSync(new Uint8Array(await workbook.xlsx.writeBuffer()))
  files['xl/charts/chart1.xml'] = strToU8(chartXml('lineChart', goodSeries))
  files['xl/charts/chart2.xml'] = strToU8(chartXml('barChart', badSeries))
  files['xl/charts/chart3.xml'] = strToU8(chartXml('lineChart', unsortedSeries))
  files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>' +
    '<Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>' +
    '<Relationship Id="rIdChart3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>' +
    '</Relationships>',
  )
  const contentTypes = new TextDecoder().decode(files['[Content_Types].xml'])
  const chartOverrides = [
    '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    '<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    '<Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
  ].join('')
  files['[Content_Types].xml'] = strToU8(contentTypes.replace('</Types>', `${chartOverrides}</Types>`))
  return Buffer.from(zipSync(files))
}

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
  assert.ok(cells['Sales!B2'] === '100')
  assert.ok(cells['Sales!A2'].startsWith('2026-01-01'))
})
