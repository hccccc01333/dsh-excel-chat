import assert from 'node:assert/strict'
import { test } from 'node:test'
import { EXCEL_ERROR_VALUES, excelErrorPattern } from '../src/error-values.ts'
import { detectErrorValues } from '../src/patterns.ts'

/**
 * `detectErrorValues` used to carry its own hand-written regex,
 * `/#(?:REF|DIV\/0|VALUE|NAME\?|N\/A|NULL|NUM)!/g`, which demanded a trailing
 * `!`. `#N/A` ends in `A` and `#NAME?` ends in `?`, so the post-edit health
 * check passed workbooks whose lookups had all failed while `excel_find_errors`
 * reported the same cells — two tools disagreeing about one workbook.
 *
 * The list now lives in one place, and this test is what keeps a consumer from
 * drifting away from it again.
 */
test('the post-edit check recognises every value the shared list declares', () => {
  const missed = EXCEL_ERROR_VALUES.filter(
    (value) => detectErrorValues({ 'Sheet1!A1': value }).length === 0,
  )
  assert.deepEqual(missed, [], `these error values are not detected: ${missed.join(', ')}`)
})

test('the lookup and name errors that used to be missed are detected', () => {
  // Named individually so a regression names the symptom, not just a count.
  for (const value of ['#N/A', '#NAME?']) {
    const found = detectErrorValues({ 'Sheet1!A1': value })
    assert.equal(found.length, 1, `${value} must be reported`)
    assert.equal(found[0].kind, 'error-value')
    assert.match(found[0].message, new RegExp(value.replace('?', '\\?')))
  }
})

test('an error inside a longer formula result is still found', () => {
  const found = detectErrorValues({ 'Sheet1!A1': '=VLOOKUP(A2,数据!A:B,2,FALSE) #N/A' })
  assert.equal(found.length, 1)
})

test('text that merely looks like an error is not reported', () => {
  // The regex is built from the list, so these must not become matches through
  // an unescaped metacharacter or a prefix of a real value.
  for (const lookalike of ['#N', 'N/A', '#REF', '#NA', 'hello', '=SUM(A1:A2)', '合计 #1']) {
    assert.deepEqual(
      detectErrorValues({ 'Sheet1!A1': lookalike }),
      [],
      `${JSON.stringify(lookalike)} must not be reported as an error`,
    )
  }
})

test('every declared value is a well-formed error token', () => {
  for (const value of EXCEL_ERROR_VALUES) {
    assert.ok(value.startsWith('#'), `${value} must start with #`)
    assert.equal(value, value.toUpperCase(), `${value} must be upper case, as Excel writes it`)
  }
  assert.equal(new Set(EXCEL_ERROR_VALUES).size, EXCEL_ERROR_VALUES.length, 'no duplicates')
})

test('the pattern escapes metacharacters in the values it is built from', () => {
  // `#NAME?` carries a `?`; unescaped it would quantify the preceding character
  // and match things like `#NAMEE` while missing `#NAME?` itself.
  const pattern = new RegExp(excelErrorPattern())
  assert.ok(pattern.test('#NAME?'))
  assert.ok(pattern.test('#DIV/0!'))
  assert.ok(!pattern.test('#NAMEE'))
})
