import assert from 'node:assert/strict'
import { test } from 'node:test'
import { buildDependencyGraph, traceDependencies } from '../src/graph.ts'

// A chain Sheet1!A1 -> B1 -> C1 -> D1, plus a fan-out where A2 feeds B2 and C2.
const formulas = [
  { id: 'Sheet1!B1', formula: '=A1+1' },
  { id: 'Sheet1!C1', formula: '=B1+1' },
  { id: 'Sheet1!D1', formula: '=C1+1' },
  { id: 'Sheet1!B2', formula: '=A2*2' },
  { id: 'Sheet1!C2', formula: '=A2*3' },
]

// The graph keys cells canonically (upper-cased sheet + column), so results come
// back upper-cased; callers map them back for display.
test('traceDependencies follows one level by default (like a single Excel click)', () => {
  const graph = buildDependencyGraph(formulas)
  const { reached, truncated } = traceDependencies(graph, 'Sheet1!A1', 'dependents')
  assert.deepEqual(reached.map((step) => step.cell), ['SHEET1!B1'])
  assert.equal(truncated, true, 'C1 sits beyond depth 1, so the caller must know the chain continues')
})

test('traceDependencies walks the whole chain when depth is large', () => {
  const graph = buildDependencyGraph(formulas)
  const { reached, truncated } = traceDependencies(graph, 'Sheet1!A1', 'dependents', 10)
  assert.deepEqual(
    reached.map((step) => `${step.cell}@${step.depth}`),
    ['SHEET1!B1@1', 'SHEET1!C1@2', 'SHEET1!D1@3'],
  )
  assert.equal(truncated, false)
})

test('traceDependencies walks precedents and fans out breadth-first', () => {
  const graph = buildDependencyGraph(formulas)
  const precedents = traceDependencies(graph, 'Sheet1!B1', 'precedents', 1).reached
  assert.deepEqual(precedents.map((step) => step.cell), ['SHEET1!A1'])

  const dependents = traceDependencies(graph, 'Sheet1!A2', 'dependents', 1).reached
  assert.deepEqual(dependents.map((step) => step.cell).sort(), ['SHEET1!B2', 'SHEET1!C2'])
})

test('traceDependencies returns nothing for a cell with no links', () => {
  const graph = buildDependencyGraph(formulas)
  assert.equal(traceDependencies(graph, 'Sheet1!Z9', 'dependents', 5).reached.length, 0)
  assert.equal(traceDependencies(graph, 'Sheet1!Z9', 'precedents', 5).reached.length, 0)
})

test('traceDependencies terminates on a cycle instead of looping', () => {
  const graph = buildDependencyGraph([
    { id: 'Sheet1!A1', formula: '=B1' },
    { id: 'Sheet1!B1', formula: '=A1' },
  ])
  const { reached } = traceDependencies(graph, 'Sheet1!A1', 'dependents', 10)
  assert.deepEqual(reached.map((step) => step.cell), ['SHEET1!B1'], 'the origin must not reappear')
  assert.equal(graph.cycles.length, 1, 'the cycle is still reported by the graph itself')
})
