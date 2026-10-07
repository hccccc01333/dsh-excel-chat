import assert from 'node:assert/strict'
import { test } from 'node:test'
import { detectEmptyGaps, detectHardcodeBreaks } from '../src/patterns.ts'

/** A four-row column whose formula rows bracket row 2, which is what a hardcode break sits inside. */
function formulaColumn(middle: string) {
  return { 'S!D1': '=A1+B1', 'S!D2': middle, 'S!D3': '=A3+B3', 'S!D4': '=A4+B4' }
}

test('a grouped number stored as text is a hardcode break', () => {
  // Excel stores a typed `1,000` as the value 1000, so it arrives ungrouped and
  // always matched. A grouped string only shows up when the number was stored as
  // *text* — pasted from a page, imported from CSV — and that is still a value
  // where the column's formula should be, so it must be reported too.
  for (const value of ['1,000', '1,234.50', '12,345,678']) {
    const found = detectHardcodeBreaks(formulaColumn(value))
    assert.equal(found.length, 1, `${value} must be reported`)
    assert.equal(found[0].kind, 'hardcode-break')
    assert.equal(found[0].cell, 'S!D2')
  }
})

test('ungrouped and otherwise numeric forms are still recognised', () => {
  for (const value of ['1000', '1234.5', '.5', '1e5', '50%', '+7', '-7']) {
    assert.equal(detectHardcodeBreaks(formulaColumn(value)).length, 1, `${value} must be reported`)
  }
})

test('a malformed grouping is not treated as a number', () => {
  // Accepting separators must not turn every comma-separated string into a value.
  for (const value of ['1,00', '1,0000', '12,34', '1,', ',100', 'abc', 'A1,B2']) {
    assert.deepEqual(
      detectHardcodeBreaks(formulaColumn(value)),
      [],
      `${JSON.stringify(value)} must not be read as a number`,
    )
  }
})

test('a value outside the formula rows is left alone', () => {
  // The check is about a value interrupting a run of formulas, not about any
  // number that happens to share the column.
  const cells = { 'S!D1': '=A1+B1', 'S!D2': '=A2+B2', 'S!D3': '1,000' }
  assert.deepEqual(detectHardcodeBreaks(cells), [])
})

test('a subtotal-structured column is exempt', () => {
  const cells = { 'S!D1': '=SUBTOTAL(9,A1:A2)', 'S!D2': '1,000', 'S!D3': '=A3+B3' }
  assert.deepEqual(detectHardcodeBreaks(cells), [])
})

test('an empty gap between formula rows is reported, a filled one is not', () => {
  const gapped = detectEmptyGaps({ 'S!A1': '=B1', 'S!A3': '=B3' })
  assert.equal(gapped.length, 1)
  assert.equal(gapped[0].kind, 'empty-gap')
  assert.equal(gapped[0].cell, 'S!A2')

  const filled = detectEmptyGaps({ 'S!A1': '=B1', 'S!A2': '5', 'S!A3': '=B3' })
  assert.deepEqual(filled, [])
})
