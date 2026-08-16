import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { access, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { applyInPlaceEdit, revertInPlaceEdit } from '../src/live-edit.ts'
import { readWorkbookCells } from '../src/workbook.ts'

async function makeWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-live-edit-'))
  const path = join(dir, 'book.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '产品'
  sheet.getCell('B1').value = '数量'
  sheet.getCell('A2').value = '苹果'
  sheet.getCell('B2').value = 10
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('applyInPlaceEdit writes the local file, backs it up, and logs the patch', async () => {
  const path = await makeWorkbook()
  const result = await applyInPlaceEdit(path, '订单!B2', 20)
  assert.equal(result.anomalies, 0)
  await access(result.backupPath)
  await access(result.patchLog)
  const cells = await readWorkbookCells(await readFile(path))
  assert.equal(cells['订单!B2'], '20')
})

test('revertInPlaceEdit restores the original local file', async () => {
  const path = await makeWorkbook()
  await applyInPlaceEdit(path, '订单!B2', 20)
  const outcome = await revertInPlaceEdit(path)
  assert.equal(outcome.restored, true)
  const cells = await readWorkbookCells(await readFile(path))
  assert.equal(cells['订单!B2'], '10')
})
