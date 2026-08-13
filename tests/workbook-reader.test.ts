import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
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
