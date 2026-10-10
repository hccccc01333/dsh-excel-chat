import assert from 'node:assert/strict'
import { t } from '../src/i18n.ts'
import { test } from 'node:test'
import { buildDependencyGraph, traceDependencies } from '../src/graph.ts'
import { validate } from '../src/validator.ts'

test('a formula that reads its own cell is a cycle', () => {
  // `A1 = A1+1` is Excel's most common circular reference — usually a typo for
  // the row above. It has to be reported, and it only is if the self-edge is
  // kept, because that edge is what the cycle detector walks.
  const graph = buildDependencyGraph([{ id: 'SHEET1!A1', formula: '=A1+1' }])
  assert.deepEqual(graph.edges, [{ from: 'SHEET1!A1', to: 'SHEET1!A1' }])
  assert.deepEqual(graph.cycles, [['SHEET1!A1', 'SHEET1!A1']])
})

test('a formula that sums a range containing itself is a cycle', () => {
  const graph = buildDependencyGraph([{ id: 'SHEET1!A1', formula: '=SUM(A1:A5)' }])
  assert.equal(graph.cycles.length, 1)
  assert.deepEqual(graph.cycles[0], ['SHEET1!A1', 'SHEET1!A1'])
})

test('mutual references are reported once', () => {
  const graph = buildDependencyGraph([
    { id: 'SHEET1!A1', formula: '=B1' },
    { id: 'SHEET1!B1', formula: '=A1' },
  ])
  assert.equal(graph.cycles.length, 1)
  assert.deepEqual([...graph.cycles[0]!].sort(), ['SHEET1!A1', 'SHEET1!A1', 'SHEET1!B1'])
})

test('a ring of three is reported in traversal order', () => {
  const graph = buildDependencyGraph([
    { id: 'SHEET1!A1', formula: '=B1' },
    { id: 'SHEET1!B1', formula: '=C1' },
    { id: 'SHEET1!C1', formula: '=A1' },
  ])
  assert.equal(graph.cycles.length, 1)
  assert.deepEqual(graph.cycles[0], ['SHEET1!A1', 'SHEET1!B1', 'SHEET1!C1', 'SHEET1!A1'])
})

test('an acyclic workbook reports no cycles', () => {
  const graph = buildDependencyGraph([
    { id: 'SHEET1!A1', formula: '=1' },
    { id: 'SHEET1!B1', formula: '=A1*2' },
    { id: 'SHEET1!C1', formula: '=A1+B1' },
  ])
  assert.deepEqual(graph.cycles, [])
  assert.deepEqual(graph.successors['SHEET1!C1'], ['SHEET1!A1', 'SHEET1!B1'])
  assert.deepEqual(graph.predecessors['SHEET1!A1'], ['SHEET1!B1', 'SHEET1!C1'])
})

test('a range reference links every cell it covers', () => {
  const graph = buildDependencyGraph([{ id: 'SHEET1!D1', formula: '=SUM(A1:B2)' }])
  assert.deepEqual(
    graph.edges.map((edge) => edge.to).sort(),
    ['SHEET1!A1', 'SHEET1!A2', 'SHEET1!B1', 'SHEET1!B2'],
  )
})

test('a range past the enumeration cap still links its corners', () => {
  // Enumerating 20 000 cells would be wasteful, so the interior is skipped — but
  // the corners must survive, otherwise the formula looks like it depends on
  // nothing at all.
  const graph = buildDependencyGraph([{ id: 'SHEET1!D1', formula: '=SUM(A1:A20000)' }])
  assert.deepEqual(graph.edges.map((edge) => edge.to).sort(), ['SHEET1!A1', 'SHEET1!A20000'])
})

test('repeating a reference produces a single edge', () => {
  const graph = buildDependencyGraph([{ id: 'SHEET1!C1', formula: '=A1+A1+A1' }])
  assert.deepEqual(graph.edges, [{ from: 'SHEET1!C1', to: 'SHEET1!A1' }])
})

test('cross-sheet references stay on their own sheet', () => {
  const graph = buildDependencyGraph([{ id: 'SHEET1!A1', formula: '=Data!B2' }])
  assert.deepEqual(graph.edges, [{ from: 'SHEET1!A1', to: 'DATA!B2' }])
})

test('unparsable cell ids and formulas are skipped, not thrown', () => {
  const graph = buildDependencyGraph([
    { id: 'not-a-cell', formula: '=A1' },
    { id: 'SHEET1!A1', formula: '=@@@' },
    { id: 'SHEET1!B1', formula: '=A1' },
  ])
  assert.deepEqual(graph.edges, [{ from: 'SHEET1!B1', to: 'SHEET1!A1' }])
})

test('trace walks precedents and dependents in opposite directions', () => {
  const graph = buildDependencyGraph([
    { id: 'SHEET1!A1', formula: '=1' },
    { id: 'SHEET1!B1', formula: '=A1' },
    { id: 'SHEET1!C1', formula: '=B1' },
  ])
  assert.deepEqual(traceDependencies(graph, 'SHEET1!B1', 'precedents', 1).reached, [
    { cell: 'SHEET1!A1', depth: 1 },
  ])
  assert.deepEqual(traceDependencies(graph, 'SHEET1!B1', 'dependents', 1).reached, [
    { cell: 'SHEET1!C1', depth: 1 },
  ])
  assert.deepEqual(traceDependencies(graph, 'SHEET1!A1', 'dependents', 2).reached, [
    { cell: 'SHEET1!B1', depth: 1 },
    { cell: 'SHEET1!C1', depth: 2 },
  ])
})

test('trace stops at the requested depth and says so', () => {
  const graph = buildDependencyGraph([
    { id: 'SHEET1!A1', formula: '=1' },
    { id: 'SHEET1!B1', formula: '=A1' },
    { id: 'SHEET1!C1', formula: '=B1' },
  ])
  assert.equal(traceDependencies(graph, 'SHEET1!C1', 'precedents', 1).truncated, true)
  assert.equal(traceDependencies(graph, 'SHEET1!C1', 'precedents', 3).truncated, false)
})

test('tracing a self-referencing cell terminates', () => {
  // The self-edge the cycle detector needs must not make the walk loop forever.
  const graph = buildDependencyGraph([{ id: 'SHEET1!A1', formula: '=A1+1' }])
  assert.deepEqual(traceDependencies(graph, 'SHEET1!A1', 'precedents', 5).reached, [])
  assert.deepEqual(traceDependencies(graph, 'SHEET1!A1', 'dependents', 5).reached, [])
})

test('the validator reports a self-reference as a circular-reference anomaly', () => {
  // End-to-end form of the first test: the cycle list is only useful if the
  // validator turns it into something the user is told about.
  const result = validate({ 'SHEET1!A1': '=A1+1', 'SHEET1!B1': '=A1' })
  const circular = result.anomalies.filter((anomaly) => anomaly.kind === 'circular-reference')
  assert.equal(circular.length, 1)
  assert.equal(circular[0]!.cell, 'SHEET1!A1')
  assert.equal(circular[0]!.confidence, 1)
  assert.equal(circular[0]!.message, t('循环引用：{cycle}', { cycle: 'SHEET1!A1 -> SHEET1!A1' }))
})
