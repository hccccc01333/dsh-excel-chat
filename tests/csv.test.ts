import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import ExcelJS from 'exceljs'
import { guardFormulaInjection, parseCsv, stringifyCsv, unguardFormulaInjection } from '../src/csv.ts'
import { operateWorkbookFile } from '../src/operations.ts'
import { readWorkbookCells } from '../src/workbook.ts'

test('a trailing newline closes the last record instead of opening an empty one', () => {
  // Every CSV file written by a normal tool ends with a newline. Treating that
  // terminator as starting a new record gave every import a phantom last row.
  assert.deepEqual(parseCsv('a,b\n'), [['a', 'b']])
  assert.deepEqual(parseCsv('a,b'), [['a', 'b']])
  assert.deepEqual(parseCsv('a\nb\n'), [['a'], ['b']])
})

test('stringify then parse round-trips exactly', () => {
  // stringifyCsv always emits a trailing CRLF, so this only holds if parseCsv
  // ignores a closing terminator. It is the property that keeps export→import
  // from growing a row on every pass.
  for (const rows of [
    [['a', 'b'], ['c', 'd']],
    [['name', 'qty'], ['apple', '3']],
    [['']],
    [['', '']],
  ]) {
    assert.deepEqual(parseCsv(stringifyCsv(rows)), rows)
  }
})

test('an empty document has no rows', () => {
  // An empty file is not one row containing one empty cell; importCsv would
  // otherwise write a cell into the sheet for it.
  assert.deepEqual(parseCsv(''), [])
  assert.equal(stringifyCsv([]), '')
})

test('a blank line in the middle is a real record and is kept', () => {
  assert.deepEqual(parseCsv('a\n\nb'), [['a'], [''], ['b']])
  assert.deepEqual(parseCsv('a\n\n'), [['a'], ['']])
})

test('CRLF and bare CR both terminate a record', () => {
  assert.deepEqual(parseCsv('a,b\r\nc,d\r\n'), [['a', 'b'], ['c', 'd']])
  assert.deepEqual(parseCsv('a\rb'), [['a'], ['b']])
})

test('quoted fields keep delimiters, doubled quotes and newlines', () => {
  assert.deepEqual(parseCsv('"a,b",c'), [['a,b', 'c']])
  assert.deepEqual(parseCsv('"a""b"'), [['a"b']])
  assert.deepEqual(parseCsv('"line1\nline2",x'), [['line1\nline2', 'x']])
})

test('a comma-only line is two empty fields, not one', () => {
  assert.deepEqual(parseCsv(','), [['', '']])
})

test('stringify quotes only the fields that need it', () => {
  assert.equal(stringifyCsv([['a', 'b']]), 'a,b\r\n')
  assert.equal(stringifyCsv([['a,b', 'c']]), '"a,b",c\r\n')
  assert.equal(stringifyCsv([['say "hi"']]), '"say ""hi"""\r\n')
  assert.equal(stringifyCsv([['two\nlines']]), '"two\nlines"\r\n')
})

test('custom delimiters work for parsing and writing', () => {
  assert.deepEqual(parseCsv('a;b;c', ';'), [['a', 'b', 'c']])
  assert.deepEqual(parseCsv('a\tb', '\t'), [['a', 'b']])
  assert.equal(stringifyCsv([['a', 'b']], ';'), 'a;b\r\n')
})

test('formula injection is neutralized for every OWASP prefix', () => {
  // Excel strips a leading tab or carriage return before deciding whether a
  // cell is a formula, so those are attack prefixes as much as = + - @ are.
  for (const value of ['=1+1', '+1', '-5', '@x', '\t=cmd', '\r=cmd', '=cmd|\' /c calc\'!A0']) {
    assert.equal(guardFormulaInjection(value), `'${value}`, `should guard ${JSON.stringify(value)}`)
  }
})

test('ordinary values are left untouched by the injection guard', () => {
  for (const value of ['safe', '3', 'a-b', 'x@y.com', ' spaced']) {
    assert.equal(guardFormulaInjection(value), value)
  }
})

test('unguarding reverses exactly what the guard did', () => {
  // The guard is one-way unless something undoes it, and a one-way guard makes
  // export→import corrupt the value: a text cell `=1+1` came back as `'=1+1`.
  for (const value of ['=1+1', '+8613800138000', '-5', '@handle', '\t=cmd', '\r=cmd']) {
    const guarded = guardFormulaInjection(value)
    assert.deepEqual(unguardFormulaInjection(guarded), { text: value, guarded: true })
  }
})

test('unguarding leaves everything the guard would not have touched', () => {
  for (const value of ['safe', '3', 'a-b', 'x@y.com', "'hello", "''=1", '', '-5']) {
    assert.deepEqual(
      unguardFormulaInjection(value),
      { text: value, guarded: false },
      `${JSON.stringify(value)} must be left alone`,
    )
  }
})

test('importing a real CSV does not add a phantom row to the sheet', async () => {
  // End-to-end form of the first test: the unit fix is only worth anything if
  // the extra row stops reaching the worksheet. It used to, because writeContent
  // assigns '' to A4 and that materializes the row.
  const dir = await mkdtemp(join(tmpdir(), 'vera-csv-'))
  const book = join(dir, 'book.xlsx')
  const csv = join(dir, 'data.csv')
  const out = join(dir, 'out.xlsx')

  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Sheet1').getCell('A1').value = 'seed'
  await workbook.xlsx.writeFile(book)
  await writeFile(csv, 'name,qty\napple,3\npear,5\n', 'utf8')

  await operateWorkbookFile(book, [{ op: 'importCsv', file: csv, sheet: 'Sheet1' }], out)

  const result = new ExcelJS.Workbook()
  await result.xlsx.readFile(out)
  const sheet = result.getWorksheet('Sheet1')!
  assert.equal(sheet.rowCount, 3)
  assert.equal(sheet.getCell('A1').value, 'name')
  assert.equal(sheet.getCell('A3').value, 'pear')
})

test('exporting writes every cell shape as text, not as an object', async () => {
  // `String(raw)` on a cell value gave `[object Object]` for hyperlinks, rich
  // text and error cells, and a locale-and-timezone string for dates — four
  // shapes written into the CSV as garbage, unreadable by anything else.
  const dir = await mkdtemp(join(tmpdir(), 'vera-csv-export-'))
  const book = join(dir, 'book.xlsx')
  const csv = join(dir, 'out.csv')

  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = { text: '官网', hyperlink: 'https://example.com' }
  sheet.getCell('A2').value = { richText: [{ text: '富' }, { text: '文本' }] }
  sheet.getCell('A3').value = { error: '#REF!' }
  sheet.getCell('A4').value = new Date(Date.UTC(2026, 0, 15, 9, 30))
  sheet.getCell('A5').value = -5
  sheet.getCell('A6').value = '=1+1'
  sheet.getCell('A7').value = { formula: 'A5*2' }
  await workbook.xlsx.writeFile(book)

  // Take the date from the reader rather than a literal: a date cell's ISO
  // rendering depends on the machine's zone, and hardcoding it made the test
  // pass here and fail on CI, which runs in UTC.
  const shownDate = (await readWorkbookCells(await readFile(book)))['Sheet1!A4']!

  await operateWorkbookFile(book, [{ op: 'exportCsv', file: csv, sheet: 'Sheet1' }], join(dir, 't.xlsx'))
  const lines = (await readFile(csv, 'utf8')).split('\r\n').filter((line) => line !== '')

  assert.deepEqual(lines, [
    '官网',
    '富文本',
    '#REF!',
    shownDate,
    '-5',
    "'=1+1",
    '=A5*2',
  ])
})

test('exporting then importing preserves the values', async () => {
  // The guard used to be one-way, so `=1+1` went out as `'=1+1` and came back
  // with the apostrophe still attached — an export/import round-trip that
  // silently corrupted the cell.
  const dir = await mkdtemp(join(tmpdir(), 'vera-csv-roundtrip-'))
  const book = join(dir, 'book.xlsx')
  const csv = join(dir, 'out.csv')
  const back = join(dir, 'back.xlsx')

  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = '=1+1'
  sheet.getCell('A2').value = '+8613800138000'
  sheet.getCell('A3').value = -5
  sheet.getCell('A4').value = new Date(Date.UTC(2026, 0, 15, 9, 30))
  sheet.getCell('A5').value = 1234.5
  await workbook.xlsx.writeFile(book)

  await operateWorkbookFile(book, [{ op: 'exportCsv', file: csv, sheet: 'Sheet1' }], join(dir, 't.xlsx'))
  await operateWorkbookFile(book, [{ op: 'importCsv', file: csv, sheet: 'Sheet1' }], back)

  const result = new ExcelJS.Workbook()
  await result.xlsx.readFile(back)
  const out = result.getWorksheet('Sheet1')!
  assert.equal(out.getCell('A1').value, '=1+1', 'a guarded text cell must come back as text, not a formula')
  assert.equal(out.getCell('A1').formula, undefined)
  assert.equal(out.getCell('A2').value, '+8613800138000')
  assert.equal(out.getCell('A3').value, -5, 'a negative number must not be guarded')
  assert.ok(out.getCell('A4').value instanceof Date)
  assert.equal(out.getCell('A5').value, 1234.5)
})
