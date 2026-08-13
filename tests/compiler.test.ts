import { test } from 'node:test'
import assert from 'node:assert/strict'
import { compileFormula } from '../src/compiler.ts'

const salesTable = {
  sheet: 'Sales',
  columns: {
    revenue: 'B',
    cost: 'C',
    profit: 'D',
    sales: 'H',
    channel: 'C',
  },
}

test('compiles the SUMIFS example from the proposal', () => {
  const ir = {
    operation: 'aggregate',
    metric: 'sales',
    function: 'SUMIFS',
    filters: [{ column: 'channel', value_from: 'A2' }],
  }
  assert.equal(compileFormula(ir, { baseCell: 'B2', table: salesTable }), '=SUMIFS(Sales!$H:$H,Sales!$C:$C,A2)')
})

test('compiles a filter value from a column name', () => {
  const ir = {
    operation: 'aggregate',
    metric: 'sales',
    function: 'SUMIFS',
    filters: [{ column: 'channel', value_from: 'channel' }],
  }
  const table = { sheet: 'Sales', columns: { sales: 'H', channel: 'A' } }
  assert.equal(compileFormula(ir, { baseCell: 'B2', table }), '=SUMIFS(Sales!$H:$H,Sales!$A:$A,A2)')
})

test('compiles a binary arithmetic formula', () => {
  const ir = {
    operation: 'binary',
    left: { kind: 'column', column: 'revenue' },
    right: { kind: 'column', column: 'cost' },
    operator: '-',
  }
  assert.equal(compileFormula(ir, { baseCell: 'D2', table: salesTable }), '=B2-C2')
})

test('compiles a ratio formula', () => {
  const ir = {
    operation: 'ratio',
    numerator: { kind: 'column', column: 'profit' },
    denominator: { kind: 'column', column: 'revenue' },
  }
  assert.equal(compileFormula(ir, { baseCell: 'E2', table: salesTable }), '=D2/B2')
})

test('compiles constant operands', () => {
  const ir = {
    operation: 'binary',
    left: { kind: 'constant', value: 1 },
    right: { kind: 'column', column: 'cost' },
    operator: '+',
  }
  assert.equal(compileFormula(ir, { baseCell: 'D2', table: salesTable }), '=1+C2')
})

test('rejects unknown operations, functions, and columns', () => {
  assert.throws(() => compileFormula({ operation: 'magic' }, { baseCell: 'B2', table: salesTable }), /unsupported IR operation/)
  assert.throws(() => compileFormula(
    { operation: 'aggregate', metric: 'sales', function: 'VLOOKUP', filters: [] },
    { baseCell: 'B2', table: salesTable },
  ), /unsupported aggregate function/)
  assert.throws(() => compileFormula(
    { operation: 'binary', left: { kind: 'column', column: 'missing' }, right: { kind: 'column', column: 'cost' }, operator: '-' },
    { baseCell: 'D2', table: salesTable },
  ), /unknown column/)
})
