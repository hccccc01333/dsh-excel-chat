import { test } from 'node:test'
import assert from 'node:assert/strict'
import {
  columnToNumber,
  numberToColumn,
  parseCellId,
  parseFormula,
  shiftFormulaRow,
} from '../src/formula.ts'
import { validate } from '../src/validator.ts'

test('parses relative A1 references', () => {
  const parsed = parseFormula('=B2-C2')
  assert.equal(parsed.references.length, 2)
  assert.deepEqual(parsed.references[0]!.start, {
    sheet: null,
    column: 'B',
    row: 2,
    absColumn: false,
    absRow: false,
  })
  assert.deepEqual(parsed.references[1]!.start, {
    sheet: null,
    column: 'C',
    row: 2,
    absColumn: false,
    absRow: false,
  })
})

test('parses absolute and mixed references', () => {
  const parsed = parseFormula('=SUM($A$1:A$3)')
  assert.equal(parsed.references.length, 1)
  const ref = parsed.references[0]!
  assert.equal(ref.start.absColumn, true)
  assert.equal(ref.start.absRow, true)
  assert.equal(ref.end?.absColumn, false)
  assert.equal(ref.end?.absRow, true)
})

test('parses cross-sheet references and whole columns', () => {
  const parsed = parseFormula('=SUMIFS(Sales!$H:$H,Sales!$A:$A,A2)')
  assert.equal(parsed.references.length, 3)
  assert.equal(parsed.references[0]!.start.sheet, 'SALES')
  assert.equal(parsed.references[0]!.start.row, null)
  assert.equal(parsed.references[0]!.end?.column, 'H')
  assert.equal(parsed.references[1]!.start.sheet, 'SALES')
  assert.equal(parsed.references[2]!.start.column, 'A')
})

test('parses range endpoints', () => {
  const parsed = parseFormula('=SUM(B2:B99)')
  assert.equal(parsed.references.length, 1)
  assert.equal(parsed.references[0]!.start.column, 'B')
  assert.equal(parsed.references[0]!.start.row, 2)
  assert.equal(parsed.references[0]!.end?.row, 99)
})

test('does not treat function names, booleans, or string literals as references', () => {
  assert.equal(parseFormula('=LOG10(5)').references.length, 0)
  assert.equal(parseFormula('=TRUE').references.length, 0)
  assert.equal(parseFormula('=IF(A1>0,"Sheet1",B1)').references.length, 2)
})

test('column helpers round-trip', () => {
  assert.equal(columnToNumber('A'), 1)
  assert.equal(columnToNumber('Z'), 26)
  assert.equal(columnToNumber('AA'), 27)
  assert.equal(numberToColumn(1), 'A')
  assert.equal(numberToColumn(27), 'AA')
  assert.equal(numberToColumn(columnToNumber('ZZ')), 'ZZ')
})

test('parses cell ids', () => {
  assert.deepEqual(parseCellId('D4'), { sheet: 'SHEET1', column: 'D', row: 4 })
  assert.deepEqual(parseCellId('Sales!D4'), { sheet: 'SALES', column: 'D', row: 4 })
})

test('detects the classic silent reference-pattern error', () => {
  const cells = {
    D2: '=B2-C2',
    D3: '=B3-C3',
    D4: '=B4-C3',
    D5: '=B5-C5',
  }
  const result = validate(cells)
  const d4 = result.anomalies.find((anomaly) => anomaly.kind === 'reference-offset' && anomaly.cell === 'D4')
  assert.ok(d4)
  assert.match(d4!.message, /row:0/)
  assert.match(d4!.actual!, /row:-1/)
  assert.equal(d4!.confidence, 0.75)
})

test('detects hardcoded value inside a formula column', () => {
  const result = validate({ D2: '=B2-C2', D3: '40', D4: '=B4-C4' })
  const anomaly = result.anomalies.find((item) => item.kind === 'hardcode-break' && item.cell === 'D3')
  assert.ok(anomaly)
})

test('detects circular references', () => {
  const result = validate({ A1: '=B1', B1: '=A1' })
  assert.ok(result.dependencyGraph.cycles.length >= 1)
  assert.ok(result.anomalies.some((anomaly) => anomaly.kind === 'circular-reference'))
})

test('detects empty gap in a formula column', () => {
  const result = validate({ D2: '=B2-C2', D4: '=B4-C4' })
  const anomaly = result.anomalies.find((item) => item.kind === 'empty-gap' && item.cell === 'SHEET1!D3')
  assert.ok(anomaly)
})

test('accepts a consistent column', () => {
  const result = validate({ D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C4' })
  assert.equal(result.anomalies.filter((anomaly) => anomaly.kind === 'reference-offset').length, 0)
})

test('flags a column with mixed formula shapes', () => {
  const result = validate({ D2: '=B2-C2', D3: '=SUM(B3:C3)' })
  const anomaly = result.anomalies.find((item) => item.kind === 'structure-mismatch' && item.cell === 'D3')
  assert.ok(anomaly)
})

test('pattern validator works across a sheet-qualified column', () => {
  const result = validate({
    'Sheet1!D2': '=B2-C2',
    'Sheet1!D3': '=B3-C3',
    'Sheet1!D4': '=B4-C3',
  })
  assert.ok(result.anomalies.some((anomaly) => anomaly.kind === 'reference-offset' && anomaly.cell === 'Sheet1!D4'))
})

test('detects Excel error values like #REF! and #DIV/0!', () => {
  const result = validate({
    A1: '=#REF!',
    A2: '=B2/C2',
    B2: '10',
    C2: '0',
    A3: '{"error":"#DIV/0!"}',
  })
  const errors = result.anomalies.filter((anomaly) => anomaly.kind === 'error-value')
  assert.equal(errors.length, 2)
  assert.ok(errors.some((anomaly) => anomaly.cell === 'A1'))
  assert.ok(errors.some((anomaly) => anomaly.cell === 'A3'))
})

test('subtotal summary rows are not reported as pattern anomalies', () => {
  const result = validate({
    D2: '=B2-C2',
    D3: '=B3-C3',
    D4: '=SUBTOTAL(9,B2:B3)',
    D5: '=B5-C5',
  })
  assert.equal(result.anomalies.filter((anomaly) => anomaly.kind === 'reference-offset').length, 0)
  assert.equal(result.anomalies.filter((anomaly) => anomaly.kind === 'structure-mismatch').length, 0)
})

test('shiftFormulaRow shifts relative rows and preserves absolute rows', () => {
  assert.equal(shiftFormulaRow('=B3-C3', 1), '=B4-C4')
  assert.equal(shiftFormulaRow('=SUM(B3:C3)', 1), '=SUM(B4:C4)')
  assert.equal(shiftFormulaRow('=SUM($B$4:C3)', 1), '=SUM($B$4:C4)')
  assert.equal(shiftFormulaRow('=Sheet1!B3+Sheet1!$C$3', 1), '=Sheet1!B4+Sheet1!$C$3')
  assert.equal(shiftFormulaRow('=B3-C3', 0), '=B3-C3')
})
