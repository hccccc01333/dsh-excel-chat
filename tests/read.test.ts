import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { readWorkbookDetail } from '../src/read.ts'

async function makeWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-read-'))
  const path = join(dir, 'book.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = '名称'
  sheet.getCell('A1').font = { bold: true }
  sheet.getCell('B1').value = 1234.5
  sheet.getCell('B1').numFmt = '#,##0.00'
  sheet.getCell('C1').value = { formula: 'A2*2' }
  sheet.getCell('D1').value = new Date(2026, 0, 15)
  sheet.getCell('E1').value = '备注'
  sheet.getCell('E1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }
  sheet.mergeCells('E1:F1')
  sheet.getCell('A2').value = '苹果'
  sheet.getCell('A3').value = '香蕉'
  sheet.getCell('A3').dataValidation = { type: 'list', formulae: ['"苹果,香蕉"'] }
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('readWorkbookDetail reports values, formulas, types, and formats precisely', async () => {
  const path = await makeWorkbook()
  const sheets = await readWorkbookDetail(path, { sheet: 'Sheet1', range: 'A1:E3' })
  assert.equal(sheets.length, 1)
  const cells = sheets[0]!.cells
  const byId = new Map(cells.map((cell) => [cell.id, cell]))
  assert.equal(byId.get('Sheet1!A1')!.value, '名称')
  assert.equal(byId.get('Sheet1!A1')!.bold, true)
  assert.equal(byId.get('Sheet1!B1')!.value, 1234.5)
  assert.equal(byId.get('Sheet1!B1')!.type, 'number')
  assert.equal(byId.get('Sheet1!B1')!.numberFormat, '#,##0.00')
  assert.equal(byId.get('Sheet1!C1')!.formula, '=A2*2')
  assert.equal(byId.get('Sheet1!C1')!.type, 'formula')
  assert.equal(byId.get('Sheet1!D1')!.type, 'date')
  assert.equal(byId.get('Sheet1!E1')!.fill, 'FFFF00')
  assert.equal(byId.get('Sheet1!E1')!.mergedTo, 'E1:F1')
  assert.equal(byId.get('Sheet1!A3')!.dataValidationType, 'list')
})

test('readWorkbookDetail can read a single exact cell', async () => {
  const path = await makeWorkbook()
  const sheets = await readWorkbookDetail(path, { cells: ['B1'] })
  assert.equal(sheets[0]!.cells.length, 1)
  assert.equal(sheets[0]!.cells[0]!.id, 'Sheet1!B1')
})

test('readWorkbookDetail caps rows with maxRows and marks the result truncated', async () => {
  const path = await makeWorkbook()
  const sheets = await readWorkbookDetail(path, { sheet: 'Sheet1', range: 'A1:E30', maxRows: 2 })
  assert.equal(sheets.length, 1)
  const result = sheets[0]!
  assert.equal(result.range, 'A1:E2')
  assert.equal(result.truncated, true)
  assert.equal(result.cells.length, 6)
  const next = await readWorkbookDetail(path, { sheet: 'Sheet1', range: 'A3:E30', maxRows: 2 })
  assert.equal(next[0]!.range, 'A3:E4')
})
