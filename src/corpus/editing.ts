import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { buildCorpusWorkbook } from './helpers.ts'

const order = (rows: Array<Array<string | number | null>>) => ({
  name: '订单',
  headers: ['产品', '数量'],
  rows,
})

export const editingTasks: FileBenchmarkTask[] = [
  {
    id: 'clean-dedupe',
    category: 'editing',
    name: '按产品去重',
    description: '订单表按产品列去重，保留第一次出现。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'clean-dedupe', [order([
      ['苹果', 10], ['苹果', 10], ['香蕉', 5], ['梨', 3],
    ])]),
    operations: [{ op: 'dedupeRows', sheet: '订单', keep: 'first' }],
    checks: [
      { id: '订单!A2', expect: '苹果' },
      { id: '订单!A3', expect: '香蕉' },
      { id: '订单!A4', expect: '梨' },
      { id: '订单!A5', expect: null },
      { id: '订单!B3', expect: '5' },
    ],
  },
  {
    id: 'clean-fill-missing-value',
    category: 'editing',
    name: '空值填固定值',
    description: '数量列空值填 0。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'clean-fill-missing-value', [order([
      ['苹果', 10], ['香蕉', null], ['梨', 3],
    ])]),
    operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 }],
    checks: [
      { id: '订单!B3', expect: '0' },
      { id: '订单!B4', expect: '3' },
    ],
  },
  {
    id: 'clean-fill-missing-forward',
    category: 'editing',
    name: '空值向上填充',
    description: '数量列空值取上一行非空值。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'clean-fill-missing-forward', [order([
      ['苹果', 10], ['香蕉', null], ['梨', 3],
    ])]),
    operations: [{ op: 'fillMissing', range: '订单!A2:B4', mode: 'forward' }],
    checks: [
      { id: '订单!B3', expect: '10' },
      { id: '订单!B4', expect: '3' },
    ],
  },
  {
    id: 'clean-remove-empty-rows',
    category: 'editing',
    name: '删除整空行',
    description: '删除完全为空的数据行。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'clean-remove-empty-rows', [order([
      ['苹果', 10], [null, null], ['梨', 3],
    ])]),
    operations: [{ op: 'removeEmptyRows', range: '订单!A2:B4' }],
    checks: [
      { id: '订单!A3', expect: '梨' },
      { id: '订单!A4', expect: null },
    ],
  },
  {
    id: 'clean-trim-text',
    category: 'editing',
    name: '去首尾空格',
    description: '产品名去掉首尾空格。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'clean-trim-text', [order([
      ['  苹果 ', 10], ['香蕉', 5],
    ])]),
    operations: [{ op: 'trimText', range: '订单!A2:A3' }],
    checks: [{ id: '订单!A2', expect: '苹果' }],
  },
  {
    id: 'clean-normalize-text',
    category: 'editing',
    name: '全角转半角',
    description: '把全角字符和全角空格标准化为半角。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'clean-normalize-text', [order([
      ['ＡＢＣ　１２３', 10],
    ])]),
    operations: [{ op: 'normalizeText', range: '订单!A2:A2' }],
    checks: [{ id: '订单!A2', expect: 'ABC 123' }],
  },
  {
    id: 'clean-split-column',
    category: 'editing',
    name: 'SKU 分列',
    description: '把 SKU 列按 - 拆成两列，原数据右移。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'clean-split-column', [{
      name: '订单',
      headers: ['SKU', '数量'],
      rows: [['SKU-01', 10], ['SKU-02', 20]],
    }]),
    operations: [{ op: 'splitColumn', sheet: '订单', column: 'A', delimiter: '-', startRow: 2, endRow: 3 }],
    checks: [
      { id: '订单!A2', expect: 'SKU' },
      { id: '订单!B2', expect: '01' },
      { id: '订单!C2', expect: '10' },
    ],
  },
  {
    id: 'clean-fuzzy-match',
    category: 'editing',
    name: '两表模糊匹配',
    description: '订单名称与价目表按相似度匹配（容忍空格/错字），回填编码。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'clean-fuzzy-match', [
      order([['苹果', 10], ['苹 果', 5], ['香蕉', 8]]),
      { name: '价目表', headers: ['名称', '编码'], rows: [['苹果', 'P01'], ['香蕉', 'P02']] },
    ]),
    operations: [{
      op: 'fuzzyMatch',
      source: '订单!A2:B4',
      sourceKey: 'A',
      target: '价目表!A2:B3',
      targetKey: 'A',
      valueColumn: 'B',
      outputColumn: 'C',
      threshold: 0.6,
    }],
    checks: [
      { id: '订单!C2', expect: 'P01' },
      { id: '订单!C3', expect: 'P01' },
      { id: '订单!C4', expect: 'P02' },
    ],
  },
  {
    id: 'format-header-bold',
    category: 'editing',
    name: '表头加粗',
    description: '表头行加粗。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'format-header-bold', [order([
      ['苹果', 10], ['香蕉', 5],
    ])]),
    operations: [{ op: 'style', range: '订单!A1:B1', style: { bold: true } }],
    checks: [{ id: '订单!A1', bold: true }],
  },
  {
    id: 'format-number-format',
    category: 'editing',
    name: '金额数字格式',
    description: '数量列应用千分位两位小数格式。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'format-number-format', [order([
      ['苹果', 1000], ['香蕉', 2500],
    ])]),
    operations: [{ op: 'style', range: '订单!B2:B3', style: { numberFormat: '#,##0.00' } }],
    checks: [{ id: '订单!B2', numberFormat: '#,##0.00' }],
  },
  {
    id: 'format-fill-color',
    category: 'editing',
    name: '单元格填充色',
    description: '重点行填充黄色。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'format-fill-color', [order([
      ['苹果', 10], ['香蕉', 5],
    ])]),
    operations: [{ op: 'style', range: '订单!A2:B2', style: { fill: 'FFFF00' } }],
    checks: [{ id: '订单!A2', fill: 'FFFF00' }],
  },
  {
    id: 'format-wrap-align',
    category: 'editing',
    name: '换行与居中',
    description: '表头自动换行并水平居中。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'format-wrap-align', [order([
      ['苹果', 10],
    ])]),
    operations: [{ op: 'style', range: '订单!A1:B1', style: { wrapText: true, hAlign: 'center' } }],
    checks: [
      { id: '订单!A1', wrapText: true },
      { id: '订单!B1', hAlign: 'center' },
    ],
  },
  {
    id: 'edit-sort-one-key',
    category: 'editing',
    name: '单键排序',
    description: '按金额降序排序。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-sort-one-key', [{
      name: '订单',
      headers: ['产品', '金额'],
      rows: [['A', 100], ['B', 300], ['C', 200]],
    }]),
    operations: [{ op: 'sortRange', range: '订单!A1:B4', keys: [{ column: 'B', direction: 'desc' }], headerRows: 1 }],
    checks: [
      { id: '订单!A2', expect: 'B' },
      { id: '订单!A4', expect: 'A' },
    ],
  },
  {
    id: 'edit-sort-multi-key',
    category: 'editing',
    name: '多键排序',
    description: '先按区域升序，再按金额降序。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-sort-multi-key', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [['华东', 100], ['华北', 300], ['华东', 200]],
    }]),
    operations: [{
      op: 'sortRange',
      range: '订单!A1:B4',
      keys: [{ column: 'A', direction: 'asc' }, { column: 'B', direction: 'desc' }],
      headerRows: 1,
    }],
    checks: [
      { id: '订单!A2', expect: '华东' },
      { id: '订单!A3', expect: '华东' },
      { id: '订单!B3', expect: '100' },
      { id: '订单!A4', expect: '华北' },
    ],
  },
  {
    id: 'edit-filter-to-range',
    category: 'editing',
    name: '筛选到新表',
    description: '筛选华东区域的订单到新表。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-filter-to-range', [{
      name: '订单',
      headers: ['区域', '产品', '金额'],
      rows: [['华东', 'A', 100], ['华北', 'B', 200], ['华东', 'C', 300]],
    }]),
    operations: [
      { op: 'addSheet', name: '华东' },
      { op: 'filterToRange', source: '订单!A1:C4', criteria: [{ column: 'A', operator: 'eq', value: '华东' }], target: '华东!A1' },
    ],
    checks: [
      { id: '华东!A2', expect: '华东' },
      { id: '华东!C2', expect: '100' },
      { id: '华东!C3', expect: '300' },
    ],
  },
  {
    id: 'edit-find-replace',
    category: 'editing',
    name: '查找替换',
    description: '把华东替换为华东大区。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-find-replace', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [['华东', 100], ['华北', 200]],
    }]),
    operations: [{ op: 'findReplace', find: '华东', replace: '华东大区', sheet: '订单' }],
    checks: [{ id: '订单!A2', expect: '华东大区' }],
  },
  {
    id: 'edit-sheet-add',
    category: 'editing',
    name: '新建工作表',
    description: '新建汇总表并写入合计。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-sheet-add', [order([
      ['苹果', 10],
    ])]),
    operations: [
      { op: 'addSheet', name: '汇总' },
      { op: 'set', cells: { '汇总!A1': '合计' } },
    ],
    checks: [{ id: '汇总!A1', expect: '合计' }],
  },
  {
    id: 'edit-sheet-duplicate',
    category: 'editing',
    name: '复制工作表',
    description: '复制订单表并修改副本数据。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-sheet-duplicate', [order([
      ['苹果', 10], ['香蕉', 5],
    ])]),
    operations: [
      { op: 'duplicateSheet', name: '订单', newName: '订单副本' },
      { op: 'set', cells: { '订单副本!A2': '新数据' } },
    ],
    checks: [
      { id: '订单副本!A2', expect: '新数据' },
      { id: '订单副本!B2', expect: '10' },
    ],
  },
]
