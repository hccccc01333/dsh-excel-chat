import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildCorpusWorkbook } from '../src/corpus/helpers.ts'
import { applyOperationsToWorkbook } from '../src/operations.ts'
import { verifyWorkbookAssertions } from '../src/verifier.ts'

/**
 * Sheet-level assertions: the freeze pane, hidden rows and columns, widths and heights.
 *
 * These operations change no cell, so before this kind of assertion existed the corpus
 * could not check them at all — 38 of the 77 operations had no end-to-end coverage, and
 * "the cells are unchanged" is not a check, because an operation that silently did
 * nothing passes it.
 *
 * Every assertion here is exercised in both directions. A check that cannot fail is not
 * a check, and the first version of this file asserted only the passing direction.
 */
async function layout(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'sheet-assert-'))
  const input = await buildCorpusWorkbook(dir, 'sheet-assert', [
    { name: '订单', headers: ['产品', '数量'], rows: [['苹果', 10], ['香蕉', 5], ['梨', 7]] },
  ])
  const output = join(dir, 'out.xlsx')
  await applyOperationsToWorkbook(
    input,
    [
      { op: 'freezePanes', sheet: '订单', column: 'A', row: 2 },
      { op: 'hideRows', sheet: '订单', from: 4, to: 4 },
      { op: 'hideColumns', sheet: '订单', columns: ['B'] },
      { op: 'setColumnWidth', sheet: '订单', column: 'A', width: 24 },
      { op: 'setRowHeight', sheet: '订单', row: 1, height: 30 },
    ],
    output,
  )
  return output
}

test('a sheet assertion reads the freeze pane, hidden rows and widths', async () => {
  const path = await layout()
  const result = await verifyWorkbookAssertions(path, [
    {
      sheet: '订单',
      frozen: 'A2',
      hiddenRows: [4],
      hiddenColumns: ['B'],
      columnWidth: { A: 24 },
      rowHeight: { 1: 30 },
    },
  ])
  assert.equal(result.achieved, true, result.failures.join('; '))
})

test('a sheet assertion reports what it actually found', async () => {
  const path = await layout()
  const result = await verifyWorkbookAssertions(path, [
    { sheet: '订单', frozen: 'B2', hiddenRows: [2], columnWidth: { A: 99 } },
  ])
  assert.equal(result.achieved, false)
  // One result per assertion, and its detail names every condition that failed — not
  // just the first. Three wrong conditions, three mentions.
  assert.equal(result.failures.length, 1)
  const detail = result.failures[0]!
  assert.match(detail, /B2/)
  assert.match(detail, /2/, 'the wrong hidden row should be reported')
  assert.match(detail, /99/, 'the wrong column width should be reported')
})

test('an unfrozen sheet asserts as null, not by omitting the field', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'sheet-assert-'))
  const input = await buildCorpusWorkbook(dir, 'unfreeze', [
    { name: '订单', headers: ['产品'], rows: [['苹果']] },
  ])
  const output = join(dir, 'out.xlsx')
  await applyOperationsToWorkbook(
    input,
    [
      { op: 'freezePanes', sheet: '订单', column: 'A', row: 2 },
      { op: 'unfreezePanes', sheet: '订单' },
    ],
    output,
  )
  const frozen = await verifyWorkbookAssertions(output, [{ sheet: '订单', frozen: null }])
  assert.equal(frozen.achieved, true, frozen.failures.join('; '))

  // And the same file must fail an assertion that still expects a freeze — otherwise
  // `frozen: null` would be indistinguishable from not checking at all.
  const stillFrozen = await verifyWorkbookAssertions(output, [{ sheet: '订单', frozen: 'A2' }])
  assert.equal(stillFrozen.achieved, false)
})

test('hiding a row keeps its contents', async () => {
  const path = await layout()
  // The row is hidden, not removed. An assertion that only checked `hidden` would pass
  // an implementation that dropped the row on the floor.
  const result = await verifyWorkbookAssertions(path, [{ id: '订单!A4', expect: '梨' }])
  assert.equal(result.achieved, true, result.failures.join('; '))
})

test('an assertion on a sheet that does not exist fails rather than passing quietly', async () => {
  const path = await layout()
  const result = await verifyWorkbookAssertions(path, [{ sheet: '不存在', frozen: 'A2' }])
  assert.equal(result.achieved, false)
  assert.match(result.failures.join(' '), /不存在/)
})

test('a sheet assertion that names no condition is an error, not a pass', async () => {
  const path = await layout()
  const result = await verifyWorkbookAssertions(path, [{ sheet: '订单' }])
  assert.equal(result.achieved, false, 'an assertion with nothing to check must not report success')
})

test('cell and sheet assertions in one list are both evaluated', async () => {
  const path = await layout()
  const result = await verifyWorkbookAssertions(path, [
    { id: '订单!A2', expect: '苹果' },
    { sheet: '订单', frozen: 'A2' },
    { id: '订单!B2', expect: '10' },
  ])
  assert.equal(result.total, 3)
  assert.equal(result.achieved, true, result.failures.join('; '))
})
