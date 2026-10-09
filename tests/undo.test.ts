import assert from 'node:assert/strict'
import { test } from 'node:test'
import ExcelJS from 'exceljs'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { operateWorkbookFile } from '../src/operations.ts'
import { restoreSnapshot, snapshotPath } from '../src/safe-write.ts'
import { readPatchLog, rollbackPatchLog } from '../src/diff.ts'
import { revertInPlaceEdit } from '../src/live-edit.ts'

async function makeBook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-undo-'))
  const path = join(dir, 'book.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '区域'
  sheet.getCell('A2').value = '华东'
  sheet.getCell('B2').value = 100
  await workbook.xlsx.writeFile(path)
  return path
}

async function read(path: string) {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await readFile(path))
  const sheet = workbook.getWorksheet('订单')!
  const a1 = sheet.getCell('A1')
  return {
    a1: a1.value,
    bold: a1.font?.bold === true,
    fill: (a1.fill as { fgColor?: { argb?: string } } | undefined)?.fgColor?.argb ?? null,
  }
}

/**
 * Undo has to restore the file, not the values.
 *
 * Every edit rewrites the whole package, so an operation changes far more than the
 * cell values a diff can see. The audit-log replay used to be the only rollback,
 * and for a formatting-only edit it did nothing at all while reporting success —
 * the log records value diffs, and there were none.
 */
test('undo restores formatting that a value-level rollback cannot see', async () => {
  const path = await makeBook()
  const before = await read(path)
  assert.equal(before.bold, false)

  const result = await operateWorkbookFile(path, [
    { op: 'style', range: '订单!A1:B1', style: { bold: true, fill: 'FFFF00' } },
  ], path)

  const styled = await read(path)
  assert.equal(styled.bold, true, 'the edit should have applied the style')

  // The premise of the bug: a formatting-only edit produces no value diff, so the
  // audit log has nothing to replay.
  const log = await readPatchLog(result.patchLog)
  assert.equal(log.patches.length, 0, 'a style-only edit should produce no value patches')

  const { restored } = await restoreSnapshot(path)
  assert.equal(restored, true)
  assert.deepEqual(await read(path), before, 'the snapshot should restore the formatting')
})

test('a value-level rollback alone leaves the formatting behind', async () => {
  // Pins the limitation the snapshot exists to avoid, so the fallback path cannot
  // be mistaken for a complete undo if it is ever used on its own.
  const path = await makeBook()
  const result = await operateWorkbookFile(path, [
    { op: 'style', range: '订单!A1:B1', style: { bold: true } },
  ], path)
  const log = await readPatchLog(result.patchLog)
  await rollbackPatchLog(path, log, path)
  assert.equal((await read(path)).bold, true, 'the value replay cannot undo formatting')
})

test('the in-place revert restores the whole file and says so', async () => {
  const path = await makeBook()
  const before = await read(path)

  await operateWorkbookFile(path, [
    { op: 'style', range: '订单!A1:B1', style: { bold: true, fill: 'FFFF00' } },
    { op: 'set', cells: { '订单!C1': '备注' } },
  ], path)
  assert.equal((await read(path)).bold, true)

  const outcome = await revertInPlaceEdit(path)
  assert.equal(outcome.restored, true)
  assert.match(outcome.message, /快照/)
  assert.deepEqual(await read(path), before)
})

test('without a snapshot the rollback is honest about being partial', async () => {
  const path = await makeBook()
  await operateWorkbookFile(path, [{ op: 'set', cells: { '订单!A2': '华北' } }], path)

  // Simulate the case the snapshot cannot cover: nothing was overwritten, so no
  // pre-edit copy was taken.
  await rm(snapshotPath(path), { force: true })

  const outcome = await revertInPlaceEdit(path)
  assert.equal(outcome.restored, true)
  assert.match(outcome.message, /没有编辑前的快照/)
  assert.equal((await read(path)).a1, '区域')
})

test('a revert with neither snapshot nor log reports that it did nothing', async () => {
  const path = await makeBook()
  await rm(snapshotPath(path), { force: true })
  const outcome = await revertInPlaceEdit(path)
  assert.equal(outcome.restored, false)
  assert.match(outcome.message, /没有可回滚/)
})

/**
 * An in-place edit must still produce an audit trail.
 *
 * `operateWorkbookFile` used to read the "before" state *after* writing, so when
 * the output path was the input path both reads returned the edited workbook and
 * every in-place edit logged zero patches — leaving the audit log empty and the
 * patch-log rollback with nothing to replay.
 */
test('an in-place edit records what it changed', async () => {
  const path = await makeBook()
  const result = await operateWorkbookFile(path, [
    { op: 'set', cells: { '订单!A2': '华北' } },
  ], path)

  const log = await readPatchLog(result.patchLog)
  assert.equal(log.patches.length, 1, 'the in-place edit should have been logged')
  assert.equal(log.patches[0]!.id, '订单!A2')
  assert.equal(log.patches[0]!.oldValue, '华东')
  assert.equal(log.patches[0]!.newValue, '华北')
})
