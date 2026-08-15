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
]
