import assert from 'node:assert/strict'
import { test } from 'node:test'
import ExcelJS from 'exceljs'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { operateWorkbookFile } from '../src/operations.ts'
import { workbookFingerprint } from '../src/workbook.ts'
import type { ExcelOperation } from '../src/operations.ts'

async function makeWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-fingerprint-'))
  const path = join(dir, 'book.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.addRow(['名称', '金额'])
  sheet.addRow(['A', 100])
  sheet.addRow(['B', 200])
  sheet.getCell('A1').value = { text: 'A站', hyperlink: 'https://a.example' }
  await workbook.xlsx.writeFile(path)
  return path
}

/**
 * Operations that change nothing in the cell grid: comments, panes, widths,
 * filters, conditional formats, validations, tables and zoom all live in their
 * own part of the file.
 *
 * The fingerprint used to be a snapshot of cell values plus bold/numberFormat/
 * fill, so it saw none of these. The agent loop forces `achieved: false` when
 * the fingerprint does not move, which meant a goal-mode step that only did one
 * of these was reported as a failure even though it had worked.
 */
const OFF_GRID_OPERATIONS: Array<[string, ExcelOperation[]]> = [
  ['addComment', [{ op: 'addComment', cell: 'Sheet1!A2', text: '备注' }]],
  ['setColumnWidth', [{ op: 'setColumnWidth', sheet: 'Sheet1', column: 'B', width: 30 }]],
  ['freezePanes', [{ op: 'freezePanes', sheet: 'Sheet1', row: 1, column: 'A' }]],
  ['autoFilter', [{ op: 'autoFilter', range: 'Sheet1!A1:B3' }]],
  ['conditionalFormatting', [{
    op: 'conditionalFormatting',
    range: 'Sheet1!B2:B3',
    rules: [{ type: 'cellIs', operator: 'greaterThan', formula: '150', fill: 'FFC7CE' }],
  }]],
  ['dataValidation', [{
    op: 'dataValidation',
    range: 'Sheet1!B2:B3',
    type: 'whole',
    operator: 'greaterThan',
    formula1: '0',
  }]],
  ['addTable', [{ op: 'addTable', range: 'Sheet1!A1:B3' }]],
  ['setZoom', [{ op: 'setZoom', sheet: 'Sheet1', zoom: 150 }]],
  ['setHyperlink onto an existing link', [{ op: 'setHyperlink', cell: 'Sheet1!A1', url: 'https://b.example', text: 'B站' }]],
  ['set (on-grid control)', [{ op: 'set', cells: { 'Sheet1!B2': '150' } }]],
]

for (const [name, operations] of OFF_GRID_OPERATIONS) {
  test(`the fingerprint moves when ${name} changes the file`, async () => {
    const path = await makeWorkbook()
    const before = await workbookFingerprint(path)
    const out = join(join(path, '..'), 'out.xlsx')
    await operateWorkbookFile(path, operations, out)
    assert.notEqual(await workbookFingerprint(out), before)
  })
}

test('a step that changes nothing is reported as unchanged', async () => {
  // The counterpart: re-serializing through exceljs must not look like a change,
  // or the check would be useless in the other direction and the agent would
  // treat a model that did nothing as having succeeded.
  const path = await makeWorkbook()
  const before = await workbookFingerprint(path)
  const out = join(join(path, '..'), 'noop.xlsx')
  await operateWorkbookFile(path, [], out)
  assert.equal(await workbookFingerprint(out), before)
})

test('setting a cell to the value it already holds is not a change', async () => {
  const path = await makeWorkbook()
  const before = await workbookFingerprint(path)
  const out = join(join(path, '..'), 'same.xlsx')
  await operateWorkbookFile(path, [{ op: 'set', cells: { 'Sheet1!B2': '100' } }], out)
  assert.equal(await workbookFingerprint(out), before)
})
