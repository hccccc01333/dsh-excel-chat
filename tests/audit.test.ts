import assert from 'node:assert/strict'
import { test } from 'node:test'
import ExcelJS from 'exceljs'
import { findErrorCells } from '../src/audit.ts'

async function workbookBytes(build: (workbook: ExcelJS.Workbook) => void): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()
  build(workbook)
  return new Uint8Array(await workbook.xlsx.writeBuffer() as ArrayBuffer)
}

test('findErrorCells reports real error values together with their formulas', async () => {
  const bytes = await workbookBytes((workbook) => {
    const sheet = workbook.addWorksheet('订单')
    sheet.getCell('A1').value = 10
    sheet.getCell('B1').value = 0
    // A formula whose cached result is an error.
    sheet.getCell('C1').value = { formula: 'A1/B1', result: { error: '#DIV/0!' } }
    sheet.getCell('D1').value = { error: '#N/A' }
  })
  const scan = await findErrorCells(bytes)
  assert.equal(scan.total, 2)
  assert.deepEqual(scan.counts, { '#DIV/0!': 1, '#N/A': 1 })
  assert.deepEqual(scan.sheetsScanned, ['订单'])
  const division = scan.errorCells.find((entry) => entry.error === '#DIV/0!')
  assert.equal(division?.cell, '订单!C1')
  assert.equal(division?.formula, '=A1/B1', 'the producing formula is what makes the report actionable')
})

test('findErrorCells does not flag text that merely looks like an error', async () => {
  const bytes = await workbookBytes((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = '#N/A'
    sheet.getCell('A2').value = '#DIV/0! 待确认'
  })
  const scan = await findErrorCells(bytes)
  assert.equal(scan.total, 0, 'a literal "#N/A" typed by a user is not an Excel error')
})

test('findErrorCells can be limited to a single sheet', async () => {
  const bytes = await workbookBytes((workbook) => {
    workbook.addWorksheet('A').getCell('A1').value = { error: '#REF!' }
    workbook.addWorksheet('B').getCell('A1').value = { error: '#REF!' }
  })
  const scan = await findErrorCells(bytes, 'B')
  assert.equal(scan.total, 1)
  assert.deepEqual(scan.sheetsScanned, ['B'])
})
