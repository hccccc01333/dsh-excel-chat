import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { explainFormula, readCellContent } from '../src/explain.ts'

test('explainFormula describes SUMIFS with ranges and plain-language function', () => {
  const explanation = explainFormula('=SUMIFS(C2:C10,A2:A10,"华东",B2:B10,">0")')
  assert.ok(explanation.summary.includes('SUMIFS'))
  const details = explanation.details.join('\n')
  assert.ok(details.includes('按多个条件求和'))
  assert.ok(details.includes('C2:C10'))
  assert.ok(explanation.references.length >= 2)
})

test('explainFormula reports VLOOKUP with cross-sheet references', () => {
  const explanation = explainFormula('=VLOOKUP(A2,Sheet2!$A$1:$B$100,2,FALSE)')
  assert.ok(explanation.summary.includes('VLOOKUP'))
  assert.ok(explanation.details.join('\n').includes('跨表引用'))
  assert.ok(explanation.references.some((ref) => ref.toLowerCase().includes('sheet2')))
})

test('explainFormula handles IF and reads a formula cell from a workbook', async () => {
  const explanation = explainFormula('=IF(D2>0,"达标","未达标")')
  assert.ok(explanation.summary.includes('IF'))
  const dir = await mkdtemp(join(tmpdir(), 'vera-explain-'))
  const path = join(dir, 'book.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('D4').value = { formula: 'B4*C4' }
  await writeFile(path, await workbook.xlsx.writeBuffer())
  assert.equal(await readCellContent(path, 'Sheet1!D4'), '=B4*C4')
  assert.equal(explainFormula(await readCellContent(path, 'D4')).formula, '=B4*C4')
})
