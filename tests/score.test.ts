import { test } from 'node:test'
import assert from 'node:assert/strict'
import { cellValueEquals, scoreWorkbookAgainstOracle } from '../src/score.ts'

test('identical cell maps pass with accuracy 1', () => {
  const cells = {
    'Sales!B2': '10',
    'Sales!D3': '=B3-C3',
  }
  const score = scoreWorkbookAgainstOracle(cells, { ...cells })
  assert.equal(score.passes, true)
  assert.equal(score.accuracy, 1)
  assert.equal(score.mismatches.length, 0)
})

test('formula case and whitespace differences are tolerated', () => {
  const candidate = { 'Sheet1!D4': '= sum( B4 : C4 )' }
  const oracle = { 'Sheet1!D4': '=SUM(B4:C4)' }
  assert.equal(cellValueEquals(candidate['Sheet1!D4']!, oracle['Sheet1!D4']!), true)
  assert.equal(scoreWorkbookAgainstOracle(candidate, oracle).passes, true)
})

test('numeric formatting differences are tolerated', () => {
  assert.equal(cellValueEquals('40', '40.0'), true)
  assert.equal(cellValueEquals('40', '41'), false)
})

test('reports changed, removed, and added cells against the oracle', () => {
  const candidate = {
    'Sheet1!B2': '10',
    'Sheet1!D4': '=B4-C4',
    'Sheet1!E9': 'extra',
  }
  const oracle = {
    'Sheet1!B2': '11',
    'Sheet1!D4': '=B4-C4',
    'Sheet1!C3': 'missing',
  }
  const score = scoreWorkbookAgainstOracle(candidate, oracle)
  assert.equal(score.passes, false)
  assert.equal(score.total, 4)
  assert.equal(score.matched, 1)
  assert.equal(score.mismatched, 3)
  const kinds = score.mismatches.map((entry) => entry.kind).sort()
  assert.deepEqual(kinds, ['added', 'changed', 'removed'])
})

test('cell ids are compared case-insensitively per sheet and cell', () => {
  const candidate = { 'sales!d4': '=B4-C4' }
  const oracle = { 'Sales!D4': '=B4-C4' }
  assert.equal(scoreWorkbookAgainstOracle(candidate, oracle).passes, true)
})

test('empty workbooks pass', () => {
  assert.equal(scoreWorkbookAgainstOracle({}, {}).passes, true)
})
