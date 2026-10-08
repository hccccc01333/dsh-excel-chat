import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import ExcelJS from 'exceljs'
import { guardFormulaInjection, parseCsv, stringifyCsv } from '../src/csv.ts'
import { operateWorkbookFile } from '../src/operations.ts'

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
