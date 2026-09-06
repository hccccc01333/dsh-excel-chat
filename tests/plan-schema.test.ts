import { test } from 'node:test'
import assert from 'node:assert/strict'
import { sanitizeAssertions, sanitizePlan } from '../src/plan-schema.ts'

test('sanitizePlan prefixes bare ranges with the first sheet', () => {
  const { steps } = sanitizePlan([{
    name: 's',
    operations: [{ op: 'style', range: 'A1:B1', style: { bold: true } }],
  }], ['订单'])
  assert.equal((steps[0]!.operations[0] as { range: string }).range, '订单!A1:B1')
})

test('sanitizePlan fills missing sheet and wraps single objects into arrays', () => {
  const { steps, notes } = sanitizePlan([{
    name: 's',
    operations: [
      { op: 'dedupeRows', columns: ['A'] },
      { op: 'sortRange', range: '订单!A1:B4', keys: { column: 'B', direction: 'desc' } },
    ],
  }], ['订单'])
  assert.equal((steps[0]!.operations[0] as { sheet: string }).sheet, '订单')
  assert.ok(Array.isArray((steps[0]!.operations[1] as { keys: unknown }).keys))
  assert.ok(notes.length >= 1)
})

test('sanitizePlan applies fillMissing aliases', () => {
  const { steps } = sanitizePlan([{
    name: 's',
    operations: [{ op: 'fillMissing', range: '订单!B2:B4', fillValue: 0 }],
  }], ['订单'])
  const op = steps[0]!.operations[0] as { mode: string; value: number }
  assert.equal(op.mode, 'value')
  assert.equal(op.value, 0)
})

test('sanitizePlan throws when a required array field is missing', () => {
  assert.throws(() => sanitizePlan([{
    name: 's',
    operations: [{ op: 'aggregateReport', source: '订单!A1:B4', groupColumn: 'A' }],
  }], ['订单']), /缺少必填数组 metrics/)
})

test('sanitizePlan requires freezePanes row and column', () => {
  assert.throws(
    () => sanitizePlan([{ operations: [{ op: 'freezePanes', sheet: '订单', range: '订单!A1:B2' }] }], ['订单']),
    /freezePanes 缺少必填数字 row/,
  )
  assert.throws(
    () => sanitizePlan([{ operations: [{ op: 'freezePanes', sheet: '订单', row: 2 }] }], ['订单']),
    /freezePanes 缺少必填字段 column/,
  )
  const fixed = sanitizePlan([{ operations: [{ op: 'freezePanes', sheet: '订单', column: 'A', row: '2' }] }], ['订单'])
  assert.equal((fixed.steps[0]!.operations[0] as Record<string, unknown>).row, 2)
})

test('sanitizePlan expands a single-cell fillSeries target to a range', () => {
  const fixed = sanitizePlan(
    [{ operations: [{ op: 'fillSeries', start: '订单!A2', target: '订单!A3' }] }],
    ['订单'],
  )
  assert.equal((fixed.steps[0]!.operations[0] as Record<string, unknown>).target, '订单!A2:A3')
})

test('sanitizePlan derives freezePanes row and column from a single-cell range', () => {
  const fixed = sanitizePlan(
    [{ operations: [{ op: 'freezePanes', sheet: '订单', range: '订单!A2' }] }],
    ['订单'],
  )
  const operation = fixed.steps[0]!.operations[0] as Record<string, unknown>
  assert.equal(operation.column, 'A')
  assert.equal(operation.row, 2)
})

test('sanitizePlan validates joinSheets arrays and key fields', () => {
  assert.throws(
    () => sanitizePlan([{
      operations: [{ op: 'joinSheets', source: '订单!A1:B4', sourceKey: 'A', lookup: '客户!A1:C3', lookupKey: 'A', valueColumns: ['B'] }],
    }], ['订单']),
    /缺少必填数组 outputColumns/,
  )
  assert.throws(
    () => sanitizePlan([{
      operations: [{ op: 'joinSheets', source: '订单!A1:B4', sourceKey: 'A', lookup: '客户!A1:C3', valueColumns: ['B'], outputColumns: ['C'] }],
    }], ['订单']),
    /缺少必填字段 lookupKey/,
  )
  const fixed = sanitizePlan([{
    operations: [{ op: 'joinSheets', source: '订单!A1:B4', sourceKey: 'A', lookup: '客户!A1:C3', lookupKey: 'A', valueColumns: ['B'], outputColumns: ['C'] }],
  }], ['订单'])
  assert.ok(fixed.steps.length === 1)
})

test('sanitizePlan normalizes crosstab flat metric fields and validates function', () => {
  const fixed = sanitizePlan([{
    operations: [{ op: 'crosstab', source: '订单!A1:C4', rowColumn: 'A', columnColumn: 'B', metricColumn: 'C' }],
  }], ['订单'])
  const op = fixed.steps[0]!.operations[0] as Record<string, unknown>
  assert.deepEqual(op.metric, { column: 'C', function: 'sum' })
  assert.throws(
    () => sanitizePlan([{
      operations: [{ op: 'crosstab', source: '订单!A1:C4', rowColumn: 'A', columnColumn: 'B', metric: { column: 'C', function: 'median' } }],
    }], ['订单']),
    /metric.function 不支持/,
  )
})

test('sanitizePlan checks rankColumn and new layout ops', () => {
  assert.throws(
    () => sanitizePlan([{ operations: [{ op: 'rankColumn', range: '订单!A1:B4', metricColumn: 'B' }] }], ['订单']),
    /rankColumn 缺少必填字段 outputColumn/,
  )
  assert.throws(
    () => sanitizePlan([{ operations: [{ op: 'hideRows', sheet: '订单', from: 2 }] }], ['订单']),
    /hideRows 缺少必填数字 to/,
  )
  const fixed = sanitizePlan([{
    operations: [{ op: 'moveSheet', name: '汇总', position: '1' }],
  }], ['订单'])
  assert.equal((fixed.steps[0]!.operations[0] as Record<string, unknown>).position, 1)
})

test('sanitizePlan salvages colloquial sheet names onto the exact sheet list', () => {
  const fixed = sanitizePlan([{
    operations: [{ op: 'renameSheet', oldName: '订单表', newName: '销售表' }],
  }], ['订单'])
  const op = fixed.steps[0]!.operations[0] as Record<string, unknown>
  assert.equal(op.oldName, '订单')
  assert.equal(op.newName, '销售表')
  const deduped = sanitizePlan([{
    operations: [{ op: 'dedupeRows', sheet: '订单表', columns: ['A'] }],
  }], ['订单'])
  assert.equal((deduped.steps[0]!.operations[0] as Record<string, unknown>).sheet, '订单')
  // Unknown names pass through untouched.
  const untouched = sanitizePlan([{
    operations: [{ op: 'renameSheet', oldName: '不存在的表', newName: 'x' }],
  }], ['订单'])
  assert.equal((untouched.steps[0]!.operations[0] as Record<string, unknown>).oldName, '不存在的表')
})

test('sanitizePlan aliases exceljs-native alignment names in style', () => {
  const fixed = sanitizePlan([{
    operations: [{ op: 'style', range: '订单!A1:B1', style: { wrapText: true, horizontal: 'center', vertical: 'middle' } }],
  }], ['订单'])
  const style = (fixed.steps[0]!.operations[0] as Record<string, unknown>).style as Record<string, unknown>
  assert.equal(style.hAlign, 'center')
  assert.equal(style.vAlign, 'middle')
})

test('sanitizeAssertions keeps well-formed assertions and repairs salvageable ones', () => {
  const { assertions, notes } = sanitizeAssertions([
    { id: '汇总!B2', startsWith: '=SUMIFS(' },
    { id: 'B3', expect: 0 },
    { id: '订单!A1', expect: null },
    { expect: 1 },
    { id: '', expect: 'x' },
    { id: '订单!C1' },
    { id: '订单!D1', startsWith: '' },
    'not-an-object',
  ], ['订单'])
  assert.deepEqual(assertions, [
    { id: '汇总!B2', startsWith: '=SUMIFS(' },
    { id: '订单!B3', expect: '0' },
    { id: '订单!A1', expect: null },
  ])
  assert.ok(notes.length >= 4)
})

test('sanitizeAssertions ignores a non-array field entirely', () => {
  const { assertions, notes } = sanitizeAssertions('nope', ['订单'])
  assert.deepEqual(assertions, [])
  assert.ok(notes.some((note) => note.includes('不是数组')))
})
