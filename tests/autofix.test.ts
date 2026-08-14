import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { access, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { autofixWorkbookFile } from '../src/autofix.ts'

async function makeBrokenWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-autofix-'))
  const path = join(dir, 'broken.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = '产品'
  sheet.getCell('B1').value = '数量'
  sheet.getCell('C1').value = '单价'
  sheet.getCell('D1').value = '金额'
  sheet.getCell('D2').value = { formula: 'B2*C2' }
  sheet.getCell('D3').value = { formula: 'B3*C3' }
  sheet.getCell('D4').value = { formula: 'B4*C9' }
  sheet.getCell('D5').value = { formula: 'B5*C5' }
  sheet.getCell('A2').value = 'A'
  sheet.getCell('B2').value = 10
  sheet.getCell('C2').value = 100
  sheet.getCell('A3').value = 'B'
  sheet.getCell('B3').value = 5
  sheet.getCell('C3').value = 200
  sheet.getCell('A4').value = 'C'
  sheet.getCell('B4').value = 8
  sheet.getCell('C4').value = 200
  sheet.getCell('A5').value = 'D'
  sheet.getCell('B5').value = 4
  sheet.getCell('C5').value = 100
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('autofixWorkbookFile repairs a silent reference-pattern error and re-validates', async () => {
  const path = await makeBrokenWorkbook()
  const outcome = await autofixWorkbookFile(path)
  assert.equal(outcome.before.total, 1)
  assert.equal(outcome.before.byKind['reference-offset'], 1)
  assert.equal(outcome.repairs.length, 1)
  assert.match(outcome.repairs[0]!.id, /D4$/)
  assert.equal(outcome.after.total, 0)
  assert.ok(outcome.message.includes('修复'))
  assert.ok(outcome.message.includes('D4'))
  assert.match(outcome.repairedPath, /\.repaired\.xlsx$/)
  await access(outcome.repairedPath)
})

test('autofixWorkbookFile writes to outPath and reports a clean workbook message', async () => {
  const path = await makeBrokenWorkbook()
  const dir = await mkdtemp(join(tmpdir(), 'vera-autofix-out-'))
  const outPath = join(dir, 'fixed.xlsx')
  const outcome = await autofixWorkbookFile(path, { outPath })
  assert.equal(outcome.repairedPath, outPath)
  await access(outPath)
  const clean = await autofixWorkbookFile(outPath)
  assert.equal(clean.before.total, 0)
  assert.ok(clean.message.includes('未发现公式异常'))
})
