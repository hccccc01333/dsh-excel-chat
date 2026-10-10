import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { buildCorpusWorkbook } from './helpers.ts'

/**
 * Work that spans two sheets: joining a lookup table, and a cross-tab.
 *
 * Both were among the operations the corpus never exercised, and both are the kind a
 * user asks for by name ("把客户名填进去", "做个区域×产品的交叉表"). Expectations were read
 * off the implementation before being written down.
 */
export const crossSheetTasks: FileBenchmarkTask[] = [
  {
    id: 'join-fill-from-lookup-sheet',
    category: 'analysis',
    name: '按客户码补客户名',
    description: '订单表按客户码，从客户表补出客户名。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'join-fill-from-lookup-sheet', [
      { name: '订单', headers: ['产品', '客户码'], rows: [['苹果', 'P01'], ['香蕉', 'P02']] },
      { name: '客户', headers: ['客户码', '客户名'], rows: [['P01', '张三'], ['P02', '李四']] },
    ]),
    operations: [{
      op: 'joinSheets',
      source: '订单!A1:B3',
      sourceKey: 'B',
      lookup: '客户!A1:B3',
      lookupKey: 'A',
      valueColumns: ['B'],
      outputColumns: ['C'],
    }],
    checks: [
      // No header is written into the output column — it stays empty, which is what
      // the implementation does. Pinned so a change here is a decision, not a drift.
      { id: '订单!C1', expect: null },
      // Both rows, so a join that only handles the first hit is not enough.
      { id: '订单!C2', expect: '张三' },
      { id: '订单!C3', expect: '李四' },
      // The lookup sheet is read, not rewritten.
      { id: '客户!A2', expect: 'P01' },
    ],
  },
  {
    id: 'analysis-crosstab-live-sumifs',
    category: 'analysis',
    name: '区域×产品交叉表',
    description: '按区域为行、产品为列生成交叉表，单元格是实时 SUMIFS。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'analysis-crosstab-live-sumifs', [{
      name: '订单',
      headers: ['区域', '产品', '金额'],
      rows: [['华东', 'A', 100], ['华东', 'B', 200], ['华北', 'A', 300]],
    }]),
    operations: [{
      op: 'crosstab',
      source: '订单!A1:C4',
      rowColumn: 'A',
      columnColumn: 'B',
      metric: { column: 'C', function: 'sum' },
      outputSheet: '交叉',
      totals: true,
    }],
    checks: [
      { id: '交叉!A1', expect: '区域\\产品' },
      { id: '交叉!B1', expect: 'A' },
      { id: '交叉!A2', expect: '华东' },
      // Live formulas, not frozen numbers: that is what makes the result update.
      { id: '交叉!B2', startsWith: '=SUMIFS(' },
      // Totals row and totals column.
      { id: '交叉!A4', expect: '总计' },
      { id: '交叉!B4', startsWith: '=SUM(' },
    ],
  },
]
