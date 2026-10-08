import assert from 'node:assert/strict'
import { test } from 'node:test'
import { canonicalCellId, columnToNumber, normalizeSheet, numberToColumn, parseCellId } from '../src/formula.ts'

test('column letters convert to their 1-based number', () => {
  assert.equal(columnToNumber('A'), 1)
  assert.equal(columnToNumber('Z'), 26)
  assert.equal(columnToNumber('AA'), 27)
  assert.equal(columnToNumber('ZZ'), 702)
  assert.equal(columnToNumber('XFD'), 16384)
})

test('column letters are case-insensitive', () => {
  assert.equal(columnToNumber('a'), 1)
  assert.equal(columnToNumber('aa'), 27)
})

test('a column that is not a letter is rejected instead of folded into a number', () => {
  // The old implementation multiplied out whatever it was given, so 'Sheet1'
  // became 229493717. That value then reached `sheet.getColumn(229493717)` in
  // groupColumns and allocated until the process ran out of memory — a model
  // passing a column *name* where a letter was expected could kill the run.
  for (const value of ['Sheet1', '', '1', 'A1', 'AAAA', 'A ', 'A-B', '$A$1']) {
    assert.throws(() => columnToNumber(value), /invalid column letter/, `should reject ${JSON.stringify(value)}`)
  }
})

test('a column past Excel\'s last column is rejected', () => {
  // 'ZZZ' is a well-formed three-letter column but it is past XFD (16384), so
  // exceljs would happily build column 18278 and produce a file Excel rejects.
  assert.equal(columnToNumber('XFD'), 16384)
  assert.throws(() => columnToNumber('XFE'), /out of range/)
  assert.throws(() => columnToNumber('ZZZ'), /out of range/)
})

test('column numbers convert back to letters', () => {
  assert.equal(numberToColumn(1), 'A')
  assert.equal(numberToColumn(26), 'Z')
  assert.equal(numberToColumn(27), 'AA')
  assert.equal(numberToColumn(16384), 'XFD')
  assert.equal(numberToColumn(0), '')
  assert.equal(numberToColumn(-1), '')
})

test('column helpers round-trip across the whole valid range', () => {
  for (const column of ['A', 'Z', 'AA', 'AZ', 'BA', 'ZZ', 'AAA', 'XFD']) {
    assert.equal(numberToColumn(columnToNumber(column)), column)
  }
})

test('sheet names are normalized to upper case without surrounding quotes', () => {
  assert.equal(normalizeSheet(null), 'SHEET1')
  assert.equal(normalizeSheet('Sheet1'), 'SHEET1')
  assert.equal(normalizeSheet("'My Sheet'"), 'MY SHEET')
  assert.equal(normalizeSheet('销售'), '销售')
})

test('canonical cell ids carry the normalized sheet and an upper-case column', () => {
  assert.equal(canonicalCellId('Sheet1', 'a', 1), 'SHEET1!A1')
  assert.equal(canonicalCellId(null, 'b', 12), 'SHEET1!B12')
})

test('cell ids parse with and without a sheet', () => {
  assert.deepEqual(parseCellId('D4'), { sheet: 'SHEET1', column: 'D', row: 4 })
  assert.deepEqual(parseCellId('Sheet1!D4'), { sheet: 'SHEET1', column: 'D', row: 4 })
  assert.deepEqual(parseCellId('Data!aa10'), { sheet: 'DATA', column: 'AA', row: 10 })
})

test('a cell id with a bad column or row is rejected', () => {
  // Row 0 and rows past Excel's 1048576 limit are as invalid as a bad column,
  // and they reach the same out-of-range allocations downstream.
  for (const id of ['A', '4', 'A0', 'A1048577', 'AAAA1', 'Sheet1!', 'Sheet1!A1:B2']) {
    assert.throws(() => parseCellId(id), /invalid cell id/, `should reject ${JSON.stringify(id)}`)
  }
  assert.deepEqual(parseCellId('A1048576'), { sheet: 'SHEET1', column: 'A', row: 1048576 })
})
