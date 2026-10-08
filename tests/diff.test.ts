import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  applyPatchLog,
  diffCellMaps,
  diffToPatches,
  diffWorkbookFiles,
  readPatchLog,
  rollbackPatchLog,
  writePatchLog,
} from '../src/diff.ts'
import { readWorkbookCells } from '../src/workbook.ts'

const originalPath = fileURLToPath(new URL('../fixtures/diff-original.xlsx', import.meta.url))
const patchedPath = fileURLToPath(new URL('../fixtures/diff-patched.xlsx', import.meta.url))
const logPath = fileURLToPath(new URL('../fixtures/diff.patch.json', import.meta.url))

async function writeWorkbook(path: string, d4: string): Promise<void> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sales')
  sheet.getCell('D2').value = { formula: 'B2-C2', result: 40 }
  sheet.getCell('D3').value = { formula: 'B3-C3', result: 80 }
  sheet.getCell('D4').value = { formula: d4.slice(1), result: 150 }
  sheet.getCell('D5').value = { formula: 'B5-C5', result: 160 }
  await workbook.xlsx.writeFile(path)
}

test('diffCellMaps reports added, removed, and changed cells', () => {
  const entries = diffCellMaps(
    { 'Sales!A1': 'x', 'Sales!B1': 'y' },
    { 'Sales!A1': 'x', 'Sales!C1': 'z' },
  )
  assert.deepEqual(entries, [
    { id: 'Sales!B1', kind: 'removed', oldValue: 'y', newValue: null },
    { id: 'Sales!C1', kind: 'added', oldValue: null, newValue: 'z' },
  ])
})

test('diffToPatches keeps only changed cells', () => {
  const entries = diffCellMaps(
    { 'Sales!D4': '=B4-C3' },
    { 'Sales!D4': '=B4-C4', 'Sales!E1': '=1' },
  )
  assert.deepEqual(diffToPatches(entries), [{
    id: 'Sales!D4',
    kind: 'formula',
    oldValue: '=B4-C3',
    newValue: '=B4-C4',
  }])
})

test('patch log apply and rollback round-trip on real files', async () => {
  await writeWorkbook(originalPath, '=B4-C3')
  await writeWorkbook(patchedPath, '=B4-C4')
  const entries = await diffWorkbookFiles(originalPath, patchedPath)
  const patches = diffToPatches(entries)
  assert.equal(patches.length, 1)
  const log = {
    version: 1 as const,
    createdAt: new Date().toISOString(),
    sourcePath: originalPath,
    patches,
  }
  await writePatchLog(logPath, log)

  await applyPatchLog(originalPath, await readPatchLog(logPath))
  const afterApply = await readWorkbookCells(await readFile(originalPath))
  assert.equal(afterApply['Sales!D4'], '=B4-C4')

  await rollbackPatchLog(originalPath, await readPatchLog(logPath))
  const afterRollback = await readWorkbookCells(await readFile(originalPath))
  assert.equal(afterRollback['Sales!D4'], '=B4-C3')
})

test('rolling a patch back restores the cell types, not just the text', async () => {
  // The log stores what `cellContent` produced, so a rollback that assigned the
  // raw string turned every restored number, boolean and date into text: undoing
  // an edit left `42` as the string "42". The round-trip above could not catch
  // it because `cellContent` renders both the same way — comparing the strings
  // showed no difference at all. Only the type tells them apart.
  const dir = await mkdtemp(join(tmpdir(), 'vera-rollback-'))
  const book = join(dir, 'book.xlsx')
  const edited = join(dir, 'edited.xlsx')
  const restored = join(dir, 'restored.xlsx')

  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = 42
  sheet.getCell('A2').value = true
  sheet.getCell('A3').value = new Date(2026, 0, 15, 9, 30)
  sheet.getCell('A4').value = { formula: 'A1*2' }
  sheet.getCell('A5').value = '文本'
  await workbook.xlsx.writeFile(book)

  // Change A1, then roll the change back.
  const editedWorkbook = new ExcelJS.Workbook()
  await editedWorkbook.xlsx.readFile(book)
  editedWorkbook.getWorksheet('Sheet1')!.getCell('A1').value = 99
  await editedWorkbook.xlsx.writeFile(edited)
  const log = {
    version: 1 as const,
    createdAt: new Date().toISOString(),
    sourcePath: book,
    patches: [{ id: 'Sheet1!A1', kind: 'value' as const, oldValue: '42', newValue: '99' }],
  }
  await writePatchLog(join(dir, 'p.patch.json'), log)
  await rollbackPatchLog(edited, await readPatchLog(join(dir, 'p.patch.json')), restored)

  const result = new ExcelJS.Workbook()
  await result.xlsx.readFile(restored)
  const out = result.getWorksheet('Sheet1')!
  const typeOf = (id: string): string => {
    const value = out.getCell(id).value
    if (value instanceof Date) return 'date'
    if (value !== null && typeof value === 'object') return 'formula' in value ? 'formula' : 'object'
    return typeof value
  }
  assert.equal(typeOf('A1'), 'number')
  assert.equal(out.getCell('A1').value, 42)
  assert.equal(typeOf('A2'), 'boolean')
  assert.equal(typeOf('A3'), 'date')
  assert.equal(typeOf('A4'), 'formula')
  assert.equal(typeOf('A5'), 'string')
})
