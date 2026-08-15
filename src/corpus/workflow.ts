import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { buildCorpusWorkbook } from './helpers.ts'

export const workflowTasks: FileBenchmarkTask[] = [
  {
    id: 'workflow-clean-then-report',
    category: 'workflow',
    name: '清洗后出报表',
    description: '先补空值、表头样式，再用 report 一步出经营报表。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-clean-then-report', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [['华东', 100], ['华东', null], ['华北', 300]],
    }]),
    operations: [
      { op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 },
      { op: 'style', range: '订单!A1:B1', style: { bold: true, fill: 'D9D9D9' } },
      { op: 'report', source: '订单!A1:B4', groupColumn: 'A', metrics: [{ column: 'B', function: 'sum' }], outputSheet: '经营报表' },
    ],
    checks: [
      { id: '订单!B3', expect: '0' },
      { id: '订单!A1', bold: true },
      { id: '经营报表!B2', startsWith: '=SUMIFS(' },
    ],
  },
  {
    id: 'workflow-clean-then-aggregate',
    category: 'workflow',
    name: '清洗后做透视',
    description: '先去空格再按区域汇总。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-clean-then-aggregate', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [[' 华东 ', 100], ['华北', 200]],
    }]),
    operations: [
      { op: 'trimText', range: '订单!A2:A3' },
      { op: 'aggregateReport', source: '订单!A1:B3', groupColumn: 'A', metrics: [{ column: 'B', function: 'sum' }], outputSheet: '汇总' },
    ],
    checks: [
      { id: '订单!A2', expect: '华东' },
      { id: '汇总!B2', startsWith: '=SUMIFS(' },
    ],
  },
  {
    id: 'workflow-fill-then-format',
    category: 'workflow',
    name: '补空值并美化',
    description: '空值向上填充，表头加粗。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-fill-then-format', [{
      name: '订单',
      headers: ['产品', '数量'],
      rows: [['苹果', 10], ['香蕉', null], ['梨', 3]],
    }]),
    operations: [
      { op: 'fillMissing', range: '订单!A2:B4', mode: 'forward' },
      { op: 'style', range: '订单!A1:B1', style: { bold: true } },
    ],
    checks: [
      { id: '订单!B3', expect: '10' },
      { id: '订单!A1', bold: true },
    ],
  },
  {
    id: 'workflow-split-then-sort',
    category: 'workflow',
    name: '分列后排序',
    description: 'SKU 分列后按数量降序。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-split-then-sort', [{
      name: '订单',
      headers: ['SKU', '数量'],
      rows: [['SKU-01', 10], ['SKU-02', 20]],
    }]),
    operations: [
      { op: 'splitColumn', sheet: '订单', column: 'A', delimiter: '-', startRow: 2, endRow: 3 },
      { op: 'sortRange', range: '订单!A1:C3', keys: [{ column: 'C', direction: 'desc' }], headerRows: 1 },
    ],
    checks: [
      { id: '订单!B2', expect: '02' },
      { id: '订单!C2', expect: '20' },
    ],
  },
  {
    id: 'workflow-clean-fill-format',
    category: 'workflow',
    name: '去重补空加样式',
    description: '先去重、再补空、最后美化表头。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-clean-fill-format', [{
      name: '订单',
      headers: ['产品', '数量'],
      rows: [['苹果', 10], ['苹果', 10], ['香蕉', null]],
    }]),
    operations: [
      { op: 'dedupeRows', sheet: '订单', keep: 'first' },
      { op: 'fillMissing', range: '订单!A2:B3', mode: 'value', value: 0 },
      { op: 'style', range: '订单!A1:B1', style: { bold: true } },
    ],
    checks: [
      { id: '订单!A3', expect: '香蕉' },
      { id: '订单!B3', expect: '0' },
      { id: '订单!A1', bold: true },
    ],
  },
  {
    id: 'workflow-trim-split-sort',
    category: 'workflow',
    name: '去空格分列排序',
    description: '先去空格、分列、再按数量排序。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-trim-split-sort', [{
      name: '订单',
      headers: ['SKU', '数量'],
      rows: [[' SKU-01 ', 10], ['SKU-02', 20]],
    }]),
    operations: [
      { op: 'trimText', range: '订单!A2:A3' },
      { op: 'splitColumn', sheet: '订单', column: 'A', delimiter: '-', startRow: 2, endRow: 3 },
      { op: 'sortRange', range: '订单!A1:C3', keys: [{ column: 'C', direction: 'desc' }], headerRows: 1 },
    ],
    checks: [
      { id: '订单!B2', expect: '02' },
      { id: '订单!C2', expect: '20' },
    ],
  },
  {
    id: 'workflow-clean-then-pivot',
    category: 'workflow',
    name: '清洗后做透视',
    description: '去重后按区域透视，再美化表头。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-clean-then-pivot', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [['华东', 100], ['华东', 100], ['华北', 200]],
    }]),
    operations: [
      { op: 'dedupeRows', sheet: '订单', keep: 'first' },
      { op: 'aggregateReport', source: '订单!A1:B3', groupColumn: 'A', metrics: [{ column: 'B', function: 'sum' }], outputSheet: '汇总' },
      { op: 'style', range: '汇总!A1:B1', style: { bold: true } },
    ],
    checks: [
      { id: '汇总!B2', startsWith: '=SUMIFS(' },
      { id: '汇总!A1', bold: true },
    ],
  },
  {
    id: 'workflow-fill-then-subtotal',
    category: 'workflow',
    name: '补空后分类汇总',
    description: '补空值后按区域小计。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-fill-then-subtotal', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [['华东', 100], ['华东', null], ['华北', 300]],
    }]),
    operations: [
      { op: 'fillMissing', range: '订单!A2:B4', mode: 'value', value: 0 },
      { op: 'subtotal', sheet: '订单', range: '订单!A1:B4', groupColumn: 'A', summaryColumns: [{ column: 'B', function: 'sum' }] },
    ],
    checks: [
      { id: '订单!A4', expect: '华东 汇总' },
      { id: '订单!B4', startsWith: '=SUBTOTAL(' },
      { id: '订单!A7', expect: '总计' },
    ],
  },
  {
    id: 'workflow-normalize-then-fuzzy',
    category: 'workflow',
    name: '标准化后模糊匹配',
    description: '文本标准化后与价目表模糊匹配。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-normalize-then-fuzzy', [
      {
        name: '订单',
        headers: ['名称', '数量'],
        rows: [['苹果', 10], ['苹 果', 5]],
      },
      { name: '价目表', headers: ['名称', '编码'], rows: [['苹果', 'P01']] },
    ]),
    operations: [
      { op: 'normalizeText', range: '订单!A2:A3' },
      { op: 'fuzzyMatch', source: '订单!A2:B3', sourceKey: 'A', target: '价目表!A2:B2', targetKey: 'A', valueColumn: 'B', outputColumn: 'C', threshold: 0.6 },
    ],
    checks: [
      { id: '订单!C2', expect: 'P01' },
      { id: '订单!C3', expect: 'P01' },
    ],
  },
  {
    id: 'workflow-dedupe-then-topn',
    category: 'workflow',
    name: '去重后取 TopN',
    description: '先去重，再按金额降序取前几。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-dedupe-then-topn', [{
      name: '订单',
      headers: ['产品', '金额'],
      rows: [['A', 100], ['A', 100], ['B', 300], ['C', 200], ['D', 50]],
    }]),
    operations: [
      { op: 'dedupeRows', sheet: '订单', keep: 'first' },
      { op: 'sortRange', range: '订单!A1:B5', keys: [{ column: 'B', direction: 'desc' }], headerRows: 1 },
      { op: 'addSheet', name: 'Top3' },
      { op: 'filterToRange', source: '订单!A1:B5', criteria: [{ column: 'B', operator: 'gte', value: 100 }], target: 'Top3!A1' },
    ],
    checks: [
      { id: 'Top3!A2', expect: 'B' },
      { id: 'Top3!A3', expect: 'C' },
      { id: 'Top3!A4', expect: 'A' },
    ],
  },
  {
    id: 'workflow-normalize-dedupe-aggregate',
    category: 'workflow',
    name: '标准化去重透视',
    description: '全角转半角、去重、再透视。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-normalize-dedupe-aggregate', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [[' 华东 ', 100], ['华东', 200], ['华北', 300]],
    }]),
    operations: [
      { op: 'normalizeText', range: '订单!A2:A4' },
      { op: 'dedupeRows', sheet: '订单', columns: ['A'], keep: 'first' },
      { op: 'aggregateReport', source: '订单!A1:B3', groupColumn: 'A', metrics: [{ column: 'B', function: 'sum' }], outputSheet: '汇总' },
    ],
    checks: [
      { id: '汇总!A2', expect: '华东' },
      { id: '汇总!A3', expect: '华北' },
      { id: '汇总!B2', startsWith: '=SUMIFS(' },
    ],
  },
  {
    id: 'workflow-sort-filter-report',
    category: 'workflow',
    name: '排序筛选再报表',
    description: '排序、筛选后按区域出报表。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-sort-filter-report', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [['华东', 100], ['华北', 200], ['华东', 300]],
    }]),
    operations: [
      { op: 'sortRange', range: '订单!A1:B4', keys: [{ column: 'B', direction: 'desc' }], headerRows: 1 },
      { op: 'addSheet', name: '大单' },
      { op: 'filterToRange', source: '订单!A1:B4', criteria: [{ column: 'B', operator: 'gte', value: 200 }], target: '大单!A1' },
      { op: 'report', source: '订单!A1:B4', groupColumn: 'A', metrics: [{ column: 'B', function: 'sum' }], outputSheet: '经营报表' },
    ],
    checks: [
      { id: '大单!A2', expect: '华东' },
      { id: '大单!B2', expect: '300' },
      { id: '经营报表!B2', startsWith: '=SUMIFS(' },
    ],
  },
  {
    id: 'workflow-highlight-then-report',
    category: 'workflow',
    name: '高亮后出报表',
    description: '先高亮重点产品再出报表。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-highlight-then-report', [{
      name: '订单',
      headers: ['产品', '金额'],
      rows: [['苹果', 100], ['香蕉', 200]],
    }]),
    operations: [
      { op: 'highlightRows', sheet: '订单', range: '订单!A1:B3', criteria: [{ column: 'A', operator: 'eq', value: '苹果' }] },
      { op: 'report', source: '订单!A1:B3', groupColumn: 'A', metrics: [{ column: 'B', function: 'sum' }], outputSheet: '经营报表' },
    ],
    checks: [
      { id: '订单!A2', fill: 'FFFF00' },
      { id: '经营报表!B2', startsWith: '=SUMIFS(' },
    ],
  },
  {
    id: 'workflow-rename-then-report',
    category: 'workflow',
    name: '改名后出报表',
    description: '工作表改名后按新表出报表。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-rename-then-report', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [['华东', 100], ['华北', 200]],
    }]),
    operations: [
      { op: 'renameSheet', oldName: '订单', newName: '销售' },
      { op: 'report', source: '销售!A1:B3', groupColumn: 'A', metrics: [{ column: 'B', function: 'sum' }], outputSheet: '经营报表' },
    ],
    checks: [{ id: '经营报表!B2', startsWith: '=SUMIFS(' }],
  },
  {
    id: 'workflow-copy-then-format',
    category: 'workflow',
    name: '复制后美化',
    description: '复制数据区并美化目标表头。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-copy-then-format', [{
      name: '订单',
      headers: ['产品', '数量'],
      rows: [['苹果', 10], ['香蕉', 20]],
    }]),
    operations: [
      { op: 'copyRange', source: '订单!A2:B3', target: '订单!D2' },
      { op: 'style', range: '订单!A1:B1', style: { bold: true } },
    ],
    checks: [
      { id: '订单!D2', expect: '苹果' },
      { id: '订单!E3', expect: '20' },
      { id: '订单!A1', bold: true },
    ],
  },
  {
    id: 'workflow-clean-then-vlookup',
    category: 'workflow',
    name: '清洗后 VLOOKUP',
    description: '去空格后用 VLOOKUP 补名称。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-clean-then-vlookup', [
      {
        name: '订单',
        headers: ['编码', '数量', '名称'],
        rows: [[' P01 ', 10], ['P02', 5]],
      },
      { name: '价目表', headers: ['编码', '名称'], rows: [['P01', '苹果'], ['P02', '香蕉']] },
    ]),
    operations: [
      { op: 'trimText', range: '订单!A2:A3' },
      { op: 'set', cells: { '订单!C2': '=VLOOKUP(A2,价目表!$A$1:$B$3,2,FALSE)', '订单!C3': '=VLOOKUP(A3,价目表!$A$1:$B$3,2,FALSE)' } },
    ],
    checks: [
      { id: '订单!A2', expect: 'P01' },
      { id: '订单!C2', startsWith: '=VLOOKUP(' },
    ],
  },
  {
    id: 'workflow-fill-then-vlookup',
    category: 'workflow',
    name: '补空后 VLOOKUP',
    description: '补全缺失编码后做 VLOOKUP。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-fill-then-vlookup', [
      {
        name: '订单',
        headers: ['编码', '数量', '名称'],
        rows: [['P01', 10], [null, 5]],
      },
      { name: '价目表', headers: ['编码', '名称'], rows: [['P01', '苹果'], ['P02', '香蕉']] },
    ]),
    operations: [
      { op: 'fillMissing', range: '订单!A2:B3', mode: 'forward' },
      { op: 'set', cells: { '订单!C2': '=VLOOKUP(A2,价目表!$A$1:$B$3,2,FALSE)', '订单!C3': '=VLOOKUP(A3,价目表!$A$1:$B$3,2,FALSE)' } },
    ],
    checks: [
      { id: '订单!A3', expect: 'P01' },
      { id: '订单!C3', startsWith: '=VLOOKUP(' },
    ],
  },
  {
    id: 'workflow-dedupe-then-subtotal',
    category: 'workflow',
    name: '去重后分类汇总',
    description: '去重后按区域小计。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workflow-dedupe-then-subtotal', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [['华东', 100], ['华东', 100], ['华北', 300]],
    }]),
    operations: [
      { op: 'dedupeRows', sheet: '订单', keep: 'first' },
      { op: 'subtotal', sheet: '订单', range: '订单!A1:B3', groupColumn: 'A', summaryColumns: [{ column: 'B', function: 'sum' }] },
    ],
    checks: [
      { id: '订单!A3', expect: '华东 汇总' },
      { id: '订单!A6', expect: '总计' },
    ],
  },
]
