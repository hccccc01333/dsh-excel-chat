import { test } from 'node:test'
import assert from 'node:assert/strict'
import { validateArgs } from '@deepseek-ai/dsh-tools'
import { formulaIrSchema } from '../src/ir-schema.ts'

const spec = { ir: { ...formulaIrSchema, required: true } }

test('formulaIrSchema accepts a valid aggregate IR', () => {
  const violations = validateArgs(spec, {
    ir: {
      operation: 'aggregate',
      metric: 'sales',
      function: 'SUMIFS',
      filters: [{ column: 'channel', value_from: 'A2' }],
    },
  })
  assert.deepEqual(violations, [])
})

test('formulaIrSchema accepts a valid binary IR', () => {
  const violations = validateArgs(spec, {
    ir: {
      operation: 'binary',
      left: { kind: 'column', column: 'revenue' },
      right: { kind: 'column', column: 'cost' },
      operator: '-',
    },
  })
  assert.deepEqual(violations, [])
})

test('formulaIrSchema rejects an unknown operation', () => {
  const violations = validateArgs(spec, { ir: { operation: 'magic' } })
  assert.ok(violations.length > 0)
})

test('formulaIrSchema rejects a malformed aggregate function', () => {
  const violations = validateArgs(spec, {
    ir: {
      operation: 'aggregate',
      metric: 'sales',
      function: 'VLOOKUP',
      filters: [],
    },
  })
  assert.ok(violations.length > 0)
})

test('formulaIrSchema rejects an operand without a kind', () => {
  const violations = validateArgs(spec, {
    ir: {
      operation: 'binary',
      left: { column: 'revenue' },
      right: { kind: 'column', column: 'cost' },
      operator: '-',
    },
  })
  assert.ok(violations.length > 0)
})
