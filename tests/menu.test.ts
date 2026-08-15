import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildWorkbookMenu } from '../src/menu.ts'

async function makeWorkbook(formulas: boolean, missing: boolean): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-menu-'))
  const path = join(dir, 'ledger.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '产品'
  sheet.getCell('B1').value = '数量'
  sheet.getCell('C1').value = '金额'
  for (let row = 2; row <= 6; row++) {
    sheet.getCell(`A${row}`).value = `产品${row}`
    sheet.getCell(`B${row}`).value = missing && row === 4 ? null : row
    sheet.getCell(`C${row}`).value = formulas ? { formula: `B${row}*10` } : row * 10
  }
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('buildWorkbookMenu summarizes the file and offers concrete next steps', async () => {
  const path = await makeWorkbook(true, true)
  const menu = await buildWorkbookMenu(path)
  assert.ok(menu.summary.includes('订单'))
  assert.ok(menu.summary.includes('表头'))
  assert.ok(menu.summary.includes('空值'))
  const ids = menu.suggestions.map((suggestion) => suggestion.id)
  for (const expected of ['fillMissing', 'clean', 'health', 'report', 'aggregate', 'pivot', 'chart', 'mail', 'preset']) {
    assert.ok(ids.includes(expected), `missing suggestion ${expected}`)
  }
  const clean = menu.suggestions.find((suggestion) => suggestion.id === 'clean')!
  assert.ok(clean.example.includes('订单'))
  assert.ok(menu.note.includes('excel_undo'))
})

test('buildWorkbookMenu skips fill/health suggestions when not relevant', async () => {
  const path = await makeWorkbook(false, false)
  const menu = await buildWorkbookMenu(path)
  const ids = menu.suggestions.map((suggestion) => suggestion.id)
  assert.ok(!ids.includes('fillMissing'))
  assert.ok(!ids.includes('health'))
  assert.ok(ids.includes('clean'))
  assert.ok(ids.includes('chart'))
})
