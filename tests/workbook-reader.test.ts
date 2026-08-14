import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { fileURLToPath } from 'node:url'
import { readWorkbookCells, validateWorkbookFile } from '../src/workbook.ts'
import { validate } from '../src/validator.ts'

const fixtureDir = fileURLToPath(new URL('../fixtures', import.meta.url))
const fixturePath = fileURLToPath(new URL('../fixtures/silent-error.xlsx', import.meta.url))

async function silentErrorWorkbook(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sales')
  sheet.getCell('A1').value = '商品'
  sheet.getCell('B1').value = 'Revenue'
  sheet.getCell('C1').value = 'Cost'
  sheet.getCell('D1').value = 'Profit'
  sheet.getCell('D2').value = { formula: 'B2-C2', result: 40 }
  sheet.getCell('D3').value = { formula: 'B3-C3', result: 80 }
  sheet.getCell('D4').value = { formula: 'B4-C3', result: 150 }
  sheet.getCell('D5').value = { formula: 'B5-C5', result: 160 }
  return (await workbook.xlsx.writeBuffer()) as Buffer
}

test('reads formulas from an in-memory workbook', async () => {
  const cells = await readWorkbookCells(await silentErrorWorkbook())
  assert.equal(cells['Sales!D2'], '=B2-C2')
  assert.equal(cells['Sales!D4'], '=B4-C3')
})

test('workbook reader feeds the validator and finds the silent error', async () => {
  const cells = await readWorkbookCells(await silentErrorWorkbook())
  const result = validate(cells)
  const anomaly = result.anomalies.find((item) => item.kind === 'reference-offset' && item.cell === 'Sales!D4')
  assert.ok(anomaly)
  assert.equal(anomaly!.actual, 'SALES col:-1 row:-1')
})

test('validateWorkbookFile reads a real .xlsx file', async () => {
  const data = await silentErrorWorkbook()
  const { mkdir, writeFile } = await import('node:fs/promises')
  await mkdir(fixtureDir, { recursive: true })
  await writeFile(fixturePath, data)
  const result = await validateWorkbookFile(fixturePath)
  assert.ok(result.anomalies.some((item) => item.kind === 'reference-offset' && item.cell === 'Sales!D4'))
})

test('reads a workbook with pivot tableParts without crashing', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = 42
  const files = unzipSync(new Uint8Array(await workbook.xlsx.writeBuffer()))
  const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml']!)
  files['xl/worksheets/sheet1.xml'] = strToU8(
    sheetXml.replace('</worksheet>', '<tableParts count="1"><tablePart r:id="rId9"/></tableParts></worksheet>'),
  )
  files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../pivotTables/pivotTable1.xml"/>' +
    '</Relationships>',
  )
  files['xl/pivotTables/pivotTable1.xml'] = strToU8('<pivotTableDefinition/>')
  const cells = await readWorkbookCells(Buffer.from(zipSync(files)))
  assert.equal(cells['Sheet1!A1'], '42')
})
