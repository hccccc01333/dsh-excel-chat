import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { profileWorkbook } from '../src/profile.ts'

async function makeWorkbook(dataRows: number): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-profile-'))
  const path = join(dir, 'ledger.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  const headers = ['名称', '数量', '单价', '金额', '日期']
  headers.forEach((header, index) => {
    sheet.getCell(`${String.fromCharCode(65 + index)}1`).value = header
  })
  for (let row = 2; row <= dataRows + 1; row++) {
    sheet.getCell(`A${row}`).value = `产品${row}`
    sheet.getCell(`B${row}`).value = row === 5 ? null : (row % 7) + 1
    sheet.getCell(`C${row}`).value = 100 + row
    sheet.getCell(`D${row}`).value = { formula: `B${row}*C${row}` }
    sheet.getCell(`E${row}`).value = new Date(2026, 0, row, 12)
  }
  const second = workbook.addWorksheet('订单')
  second.getCell('A1').value = '产品'
  second.getCell('B1').value = '金额'
  second.getCell('A2').value = 'A'
  second.getCell('B2').value = 10
  second.getCell('A3').value = 'B'
  second.getCell('B3').value = 20
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('profileWorkbook reports headers, dtypes, stats, and read hints', async () => {
  const path = await makeWorkbook(6)
  const profile = await profileWorkbook(path)
  assert.equal(profile.sheetCount, 2)
  assert.equal(profile.sheets.length, 2)
  const sheet = profile.sheets.find((entry) => entry.sheet === 'Sheet1')!
  assert.equal(sheet.headerRow, 1)
  assert.equal(sheet.formulaCells, 6)
  assert.equal(sheet.dataRows, 6)
  assert.equal(sheet.truncated, false)
  assert.equal(sheet.usedRange, 'A1:E7')
  assert.equal(sheet.readHint, 'A2:E7')
  const byColumn = new Map(sheet.columns.map((column) => [column.column, column]))
  const name = byColumn.get('A')!
  assert.equal(name.header, '名称')
  assert.equal(name.dtype, 'string')
  assert.equal(name.nonEmpty, 6)
  assert.equal(name.samples.length, 3)
  const quantity = byColumn.get('B')!
  assert.equal(quantity.dtype, 'number')
  assert.equal(quantity.nonEmpty, 5)
  assert.equal(quantity.missing, 1)
  assert.ok(quantity.min !== undefined)
  assert.ok(quantity.max !== undefined)
  const date = byColumn.get('E')!
  assert.equal(date.dtype, 'date')
  assert.equal(date.minDate, '2026-01-02')
  assert.equal(date.maxDate, '2026-01-07')
})

test('profileWorkbook readHint pages large sheets in pageSize chunks', async () => {
  const path = await makeWorkbook(150)
  const profile = await profileWorkbook(path, 'Sheet1')
  assert.equal(profile.sheets.length, 1)
  const sheet = profile.sheets[0]!
  assert.equal(sheet.dataRows, 150)
  assert.equal(sheet.readHint, 'A2:E101')
  assert.equal(profile.pageSize, 100)
})
