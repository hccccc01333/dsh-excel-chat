import { test } from 'node:test'
import assert from 'node:assert/strict'
import { detectTableFromCells } from '../src/tables.ts'

test('detects a header row from text cells and maps names to column letters', () => {
  const cells = {
    'Sales!A1': 'Quarter',
    'Sales!B2': 'Revenue',
    'Sales!C2': 'Cost',
    'Sales!B3': '100',
    'Sales!C3': '80',
    'Sales!B4': '120',
    'Sales!C4': '90',
  }
  assert.deepEqual(detectTableFromCells(cells), {
    sheet: 'Sales',
    columns: { Revenue: 'B', Cost: 'C' },
  })
})

test('returns null when no row has at least two text cells', () => {
  const cells = {
    'Sheet1!A1': 'Header only one entry',
    'Sheet1!A2': '10',
    'Sheet1!B2': '=A2*2',
  }
  assert.equal(detectTableFromCells(cells), null)
})

test('filters by sheet name and picks the first header row on that sheet', () => {
  const cells = {
    'Sales!B2': 'Revenue',
    'Sales!C2': 'Cost',
    'Marketing!D2': 'Spend',
    'Marketing!E2': 'ROAS',
  }
  assert.deepEqual(detectTableFromCells(cells, 'Marketing'), {
    sheet: 'Marketing',
    columns: { Spend: 'D', ROAS: 'E' },
  })
  assert.equal(detectTableFromCells(cells, 'Ops'), null)
})

test('skips formula and numeric cells while scanning for headers', () => {
  const cells = {
    B2: '=SUM(B3:B9)',
    C2: '123',
    B3: 'Revenue',
    C3: 'Cost',
  }
  assert.deepEqual(detectTableFromCells(cells), {
    sheet: 'Sheet1',
    columns: { Revenue: 'B', Cost: 'C' },
  })
})

test('ignores cells that are empty after trimming', () => {
  const cells = {
    'Sheet1!A1': '   ',
    'Sheet1!B1': 'X',
    'Sheet1!C1': 'Y',
  }
  assert.deepEqual(detectTableFromCells(cells), {
    sheet: 'Sheet1',
    columns: { X: 'B', Y: 'C' },
  })
})
