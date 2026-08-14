import type { BenchmarkTask } from './benchmark.ts'

const table = { sheet: 'Sheet1', columns: { revenue: 'B', cost: 'C' } }

/**
 * Pass@1 benchmark cases modeled on the silent-error scenarios VERA targets:
 * per-column reference offsets, range endpoints, absolute modifiers, missing
 * fill cells, cross-sheet and multi-sheet patterns, plus LLM-only cases.
 */
export const benchmarkTasks: BenchmarkTask[] = [
  {
    name: 'silent-offset',
    cells: { D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C3', D5: '=B5-C5' },
    oracleCells: { D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C4', D5: '=B5-C5' },
    table,
  },
  {
    name: 'range-tail',
    cells: { D2: '=SUM(B2:C2)', D3: '=SUM(B3:C3)', D4: '=SUM(B4:C3)', D5: '=SUM(B5:C5)' },
    oracleCells: { D2: '=SUM(B2:C2)', D3: '=SUM(B3:C3)', D4: '=SUM(B4:C4)', D5: '=SUM(B5:C5)' },
    table,
  },
  {
    name: 'range-both-ends',
    cells: { D2: '=SUM(B2:C2)', D3: '=SUM(B3:C3)', D4: '=SUM(B3:C3)', D5: '=SUM(B5:C5)' },
    oracleCells: { D2: '=SUM(B2:C2)', D3: '=SUM(B3:C3)', D4: '=SUM(B4:C4)', D5: '=SUM(B5:C5)' },
    table,
  },
  {
    name: 'absolute-tail',
    cells: { D2: '=SUM($B$4:C2)', D3: '=SUM($B$4:C3)', D4: '=SUM($B$4:C3)', D5: '=SUM($B$4:C5)' },
    oracleCells: { D2: '=SUM($B$4:C2)', D3: '=SUM($B$4:C3)', D4: '=SUM($B$4:C4)', D5: '=SUM($B$4:C5)' },
    table,
  },
  {
    name: 'empty-gap',
    cells: { D2: '=B2-C2', D3: '=B3-C3', D5: '=B5-C5', D6: '=B6-C6' },
    oracleCells: { D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C4', D5: '=B5-C5', D6: '=B6-C6' },
    table,
  },
  {
    name: 'cross-sheet',
    cells: {
      'Sheet2!E2': '=Sheet1!B2-Sheet1!C2',
      'Sheet2!E3': '=Sheet1!B3-Sheet1!C3',
      'Sheet2!E4': '=Sheet1!B4-Sheet1!C3',
      'Sheet2!E5': '=Sheet1!B5-Sheet1!C5',
    },
    oracleCells: {
      'Sheet2!E2': '=Sheet1!B2-Sheet1!C2',
      'Sheet2!E3': '=Sheet1!B3-Sheet1!C3',
      'Sheet2!E4': '=Sheet1!B4-Sheet1!C4',
      'Sheet2!E5': '=Sheet1!B5-Sheet1!C5',
    },
    table: { sheet: 'Sheet2', columns: { revenue: 'B', cost: 'C' } },
  },
  {
    name: 'multi-sheet',
    cells: {
      'Sales!D2': '=B2-C2',
      'Sales!D3': '=B3-C3',
      'Sales!D4': '=B4-C3',
      'Sales!D5': '=B5-C5',
      'Marketing!E2': '=B2+C2',
      'Marketing!E3': '=B3+C3',
      'Marketing!E4': '=B4+C3',
      'Marketing!E5': '=B5+C5',
    },
    oracleCells: {
      'Sales!D2': '=B2-C2',
      'Sales!D3': '=B3-C3',
      'Sales!D4': '=B4-C4',
      'Sales!D5': '=B5-C5',
      'Marketing!E2': '=B2+C2',
      'Marketing!E3': '=B3+C3',
      'Marketing!E4': '=B4+C4',
      'Marketing!E5': '=B5+C5',
    },
    table,
  },
  {
    name: 'structure-mismatch',
    cells: { D2: '=B2-C2', D3: '=SUM(B3:C3)', D4: '=B4-C4', D5: '=B5-C5' },
    oracleCells: { D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C4', D5: '=B5-C5' },
    table,
  },
  {
    name: 'hardcode-break',
    cells: { D2: '=B2-C2', D3: '=B3-C3', D4: '100', D5: '=B5-C5' },
    oracleCells: { D2: '=B2-C2', D3: '=B3-C3', D4: '=B4-C4', D5: '=B5-C5' },
    table,
  },
  {
    name: 'aggregate-mismatch',
    cells: {
      D2: '=SUM(Sheet1!$B:$B)',
      D3: '=SUM(Sheet1!$B:$B)',
      D4: '=B4-C4',
      D5: '=SUM(Sheet1!$B:$B)',
    },
    oracleCells: {
      D2: '=SUM(Sheet1!$B:$B)',
      D3: '=SUM(Sheet1!$B:$B)',
      D4: '=SUM(Sheet1!$B:$B)',
      D5: '=SUM(Sheet1!$B:$B)',
    },
    table,
  },
  {
    name: 'clean-noop',
    cells: { D2: '=B2-C2', D3: '=B3-C3' },
    oracleCells: { D2: '=B2-C2', D3: '=B3-C3' },
    table,
  },
]
