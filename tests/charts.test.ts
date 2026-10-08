import assert from 'node:assert/strict'
import { test } from 'node:test'
import { expandRange, parseChartXml, parseRangeRef } from '../src/charts.ts'

test('parses a sheet-qualified range into its corners', () => {
  assert.deepEqual(parseRangeRef('Sheet1!A1:B2'), {
    sheet: 'SHEET1',
    startColumn: 'A',
    startRow: 1,
    endColumn: 'B',
    endRow: 2,
  })
  assert.deepEqual(parseRangeRef('Sheet1!A1'), {
    sheet: 'SHEET1',
    startColumn: 'A',
    startRow: 1,
    endColumn: 'A',
    endRow: 1,
  })
})

test('sheet names and columns are normalized, absolute markers dropped', () => {
  const range = parseRangeRef("'My Sheet'!$a$1:$b$2")
  assert.deepEqual(range, {
    sheet: 'MY SHEET',
    startColumn: 'A',
    startRow: 1,
    endColumn: 'B',
    endRow: 2,
  })
})

test('a range without a sheet is rejected rather than half-parsed', () => {
  // Making the `!` optional let the greedy sheet group eat the range: `A1:B2`
  // came back as sheet "A1:" covering only B2. That silent wrong answer is worse
  // than a rejection, because the caller would go on to validate the wrong cell.
  for (const ref of ['A1:B2', 'A1', 'A1:', 'Sheet1!A1:B', 'Sheet1!1:2', '', 'garbage']) {
    assert.equal(parseRangeRef(ref), null, `should reject ${JSON.stringify(ref)}`)
  }
})

test('expandRange normalizes a reversed range and enumerates column-major', () => {
  const range = parseRangeRef('Sheet1!B2:A1')!
  assert.deepEqual(expandRange(range), ['SHEET1!A1', 'SHEET1!A2', 'SHEET1!B1', 'SHEET1!B2'])
})

test('parseChartXml reads the type and the series ranges', () => {
  const xml = [
    '<c:barChart>',
    '<c:ser><c:idx val="0"/>',
    '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>',
    '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>',
    '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val>',
    '</c:ser>',
    '</c:barChart>',
  ].join('')
  assert.deepEqual(parseChartXml(xml), {
    type: 'barChart',
    series: [{
      name: 'Sheet1!$B$1',
      categories: 'Sheet1!$A$2:$A$4',
      values: 'Sheet1!$B$2:$B$4',
    }],
  })
})

test('parseChartXml reads every series, not just the first', () => {
  const xml = '<c:lineChart><c:ser><c:val><c:numRef><c:f>Sheet1!$B$2</c:f></c:numRef></c:val></c:ser>'
    + '<c:ser><c:val><c:numRef><c:f>Sheet1!$C$2</c:f></c:numRef></c:val></c:ser></c:lineChart>'
  const parsed = parseChartXml(xml)
  assert.equal(parsed.series.length, 2)
  assert.equal(parsed.series[1]!.values, 'Sheet1!$C$2')
})

test('a chart with no series and an unknown type degrades instead of throwing', () => {
  assert.deepEqual(parseChartXml('<c:lineChart></c:lineChart>'), { type: 'lineChart', series: [] })
  // 3D variants are not in the recognised list, so the type is unknown while the
  // series still come through — callers must tolerate a null type.
  assert.equal(parseChartXml('<c:bar3DChart><c:ser></c:ser></c:bar3DChart>').type, null)
})

test('a series titled with a literal string has no reference', () => {
  const parsed = parseChartXml('<c:pieChart><c:ser><c:tx><c:v>直接写</c:v></c:tx></c:ser></c:pieChart>')
  assert.equal(parsed.series[0]!.name, undefined)
})
