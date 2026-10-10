import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { buildCorpusWorkbook } from './helpers.ts'

/**
 * Sheet view and layout: freezing panes, hiding rows and columns, widths and heights.
 *
 * These operations change no cell, so the corpus could not assert them at all — the only
 * assertion kind was `{ id, expect }`, and every one of these leaves every cell exactly
 * as it was. That is why 38 of the 77 operations had no end-to-end coverage, and why
 * "the cell does not move" is not a check: an operation that silently did nothing would
 * have passed it.
 *
 * The assertions here name the sheet instead of a cell, which is the subject these
 * operations actually change. Every one was read off the implementation first: the
 * expected freeze cell is the `topLeftCell` that `freezePanes` builds, and it is
 * `column + row`, not the cell the caller passes.
 */
const order = (rows: Array<Array<string | number | null>>) => ({
  name: '订单',
  headers: ['产品', '数量'],
  rows,
})

const twoRows = [['苹果', 10], ['香蕉', 5]]

export const viewTasks: FileBenchmarkTask[] = [
  {
    id: 'view-freeze-first-row',
    category: 'editing',
    name: '冻结首行',
    description: '冻结表头行，滚动时表头留在原地。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'view-freeze-first-row', [order(twoRows)]),
    operations: [{ op: 'freezePanes', sheet: '订单', column: 'A', row: 2 }],
    checks: [{ sheet: '订单', frozen: 'A2' }],
  },
  {
    id: 'view-freeze-row-and-column',
    category: 'editing',
    name: '同时冻结首行首列',
    description: '冻结首行与首列，得到 A2 为左上角的冻结区域。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'view-freeze-row-and-column', [order(twoRows)]),
    operations: [{ op: 'freezePanes', sheet: '订单', column: 'B', row: 2 }],
    checks: [{ sheet: '订单', frozen: 'B2' }],
  },
  {
    id: 'view-unfreeze',
    category: 'editing',
    name: '取消冻结',
    description: '先冻结再取消，文件里不该留下冻结状态。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'view-unfreeze', [order(twoRows)]),
    operations: [
      { op: 'freezePanes', sheet: '订单', column: 'A', row: 2 },
      { op: 'unfreezePanes', sheet: '订单' },
    ],
    // `null` rather than omitting the field: omitting it asserts nothing, and "no freeze"
    // is the whole content of this task.
    checks: [{ sheet: '订单', frozen: null }],
  },
  {
    id: 'view-hide-rows',
    category: 'editing',
    name: '隐藏数据行',
    description: '隐藏第 3 行，内容仍在文件里，只是不显示。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'view-hide-rows', [order([...twoRows, ['梨', 7]])]),
    // Row 1 is the header, so 梨 — the third data row — is row 4. Written as row 3
    // first, and the task failed: the header row is easy to forget when counting.
    operations: [{ op: 'hideRows', sheet: '订单', from: 4, to: 4 }],
    checks: [
      { sheet: '订单', hiddenRows: [4] },
      // The row is hidden, not deleted — the value has to survive. A task that only
      // asserted `hidden` would pass an implementation that dropped the row.
      { id: '订单!A4', expect: '梨' },
    ],
  },
  {
    id: 'view-hide-columns',
    category: 'editing',
    name: '隐藏整列',
    description: '隐藏 B 列，数据保留。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'view-hide-columns', [order(twoRows)]),
    operations: [{ op: 'hideColumns', sheet: '订单', columns: ['B'] }],
    checks: [
      { sheet: '订单', hiddenColumns: ['B'] },
      { id: '订单!B2', expect: '10' },
    ],
  },
  {
    id: 'view-column-width-and-row-height',
    category: 'editing',
    name: '设置列宽与行高',
    description: '把产品列加宽、表头行加高，长品名不再被截断。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'view-column-width-and-row-height', [order(twoRows)]),
    operations: [
      { op: 'setColumnWidth', sheet: '订单', column: 'A', width: 24 },
      { op: 'setRowHeight', sheet: '订单', row: 1, height: 30 },
    ],
    checks: [{ sheet: '订单', columnWidth: { A: 24 }, rowHeight: { 1: 30 } }],
  },
  {
    id: 'view-layout-does-not-touch-values',
    category: 'editing',
    name: '只改视图不动数据',
    description: '一次做完全部视图调整，单元格内容必须一字不改。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'view-layout-does-not-touch-values', [order([...twoRows, ['梨', 7]])]),
    operations: [
      { op: 'freezePanes', sheet: '订单', column: 'A', row: 2 },
      { op: 'hideRows', sheet: '订单', from: 4, to: 4 },
      { op: 'setColumnWidth', sheet: '订单', column: 'A', width: 24 },
    ],
    checks: [
      { sheet: '订单', frozen: 'A2', hiddenRows: [4], columnWidth: { A: 24 } },
      { id: '订单!A1', expect: '产品' },
      { id: '订单!A4', expect: '梨' },
      { id: '订单!B4', expect: '7' },
    ],
  },
]
