import { test } from 'node:test'
import assert from 'node:assert/strict'
import {
  FULLSCREEN_RENDER_COLS,
  FULLSCREEN_RENDER_ROWS,
  INLINE_RENDER_COLS,
  INLINE_RENDER_ROWS,
  cellText,
  columnName,
  measureSheets,
  parseCellId,
  toGridData,
  type TableSheetShape,
} from '../src/client/view-model.ts'

test('columnName maps 1-based indexes onto Excel letters', () => {
  assert.equal(columnName(1), 'A')
  assert.equal(columnName(26), 'Z')
  assert.equal(columnName(27), 'AA')
  assert.equal(columnName(52), 'AZ')
  assert.equal(columnName(702), 'ZZ')
  assert.equal(columnName(703), 'AAA')
})

test('cellText prefers the formula and stringifies the rest', () => {
  assert.equal(cellText(42, 'B2*C2'), 'B2*C2')
  assert.equal(cellText(null), '')
  assert.equal(cellText(undefined), '')
  assert.equal(cellText(0), '0')
  assert.equal(cellText('产品'), '产品')
  assert.equal(cellText({ richText: [] }), '{"richText":[]}')
})

test('parseCellId reads sheet-qualified A1 addresses and rejects the rest', () => {
  assert.deepEqual(parseCellId('Sheet1!A1'), { row: 1, col: 1 })
  assert.deepEqual(parseCellId('订单!AB12'), { row: 12, col: 28 })
  assert.deepEqual(parseCellId('b3'), { row: 3, col: 2 })
  assert.equal(parseCellId('Sheet1!A'), null)
  assert.equal(parseCellId('Sheet1!A0'), null)
  assert.equal(parseCellId('Sheet1!A1:B2'), null)
  assert.deepEqual(parseCellId('A1'), { row: 1, col: 1 })
})

function sheet(name: string, cells: Array<[string, unknown, string?]>): TableSheetShape {
  return {
    sheet: name,
    cells: cells.map(([id, value, formula]) => (formula === undefined ? { id, value } : { id, value, formula })),
  }
}

test('measureSheets floors tiny sheets and reports the caps it used', () => {
  const layout = measureSheets([sheet('Sheet1', [['Sheet1!A1', 'x']])], false)
  assert.equal(layout.rows, 10)
  assert.equal(layout.cols, 5)
  assert.equal(layout.capRows, INLINE_RENDER_ROWS)
  assert.equal(layout.capCols, INLINE_RENDER_COLS)
  assert.equal(layout.truncated, false)
})

test('measureSheets caps an oversized sheet and flags the truncation', () => {
  const cells: Array<[string, unknown]> = []
  cells.push([`Sheet1!${columnName(INLINE_RENDER_COLS + 1)}${INLINE_RENDER_ROWS + 1}`, 1])
  const layout = measureSheets([sheet('Sheet1', cells)], false)
  assert.equal(layout.rows, INLINE_RENDER_ROWS)
  assert.equal(layout.cols, INLINE_RENDER_COLS)
  assert.equal(layout.truncated, true)
  const full = measureSheets([sheet('Sheet1', cells)], true)
  assert.equal(full.rows, INLINE_RENDER_ROWS + 1)
  assert.equal(full.cols, INLINE_RENDER_COLS + 1)
  assert.equal(full.capRows, FULLSCREEN_RENDER_ROWS)
  assert.equal(full.capCols, FULLSCREEN_RENDER_COLS)
  assert.equal(full.truncated, false)
})

test('measureSheets takes the widest sheet, not the first', () => {
  const layout = measureSheets([
    sheet('一', [['一!A1', 'x']]),
    sheet('二', [['二!C7', 'y']]),
  ], false)
  assert.equal(layout.rows, 10)
  assert.equal(layout.cols, 5)
  const wider = measureSheets([
    sheet('一', [['一!A1', 'x']]),
    sheet('二', [['二!H20', 'y']]),
  ], false)
  assert.equal(wider.rows, 20)
  assert.equal(wider.cols, 8)
})

test('toGridData converts A1 ids into 0-based row/col keys with display text', () => {
  const data = toGridData([
    sheet('订单', [
      ['订单!A1', '产品'],
      ['订单!B2', 12],
      ['订单!C3', null, 'B3*10'],
      ['订单!broken', 'ignored'],
      ['订单!A0', 'ignored'],
    ]),
    sheet('汇总', [['汇总!A1', 'ok']]),
  ])
  assert.equal(data.length, 2)
  assert.deepEqual(data[0], {
    name: '订单',
    rows: {
      0: { cells: { 0: { text: '产品' } } },
      1: { cells: { 1: { text: '12' } } },
      2: { cells: { 2: { text: 'B3*10' } } },
    },
  })
  assert.deepEqual(data[1], { name: '汇总', rows: { 0: { cells: { 0: { text: 'ok' } } } } })
})
