import { test } from 'node:test'
import assert from 'node:assert/strict'
import { sanitizePlan } from '../src/plan-schema.ts'

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
