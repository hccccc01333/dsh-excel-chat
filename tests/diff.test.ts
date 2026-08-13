import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { readFile } from 'node:fs/promises'
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
