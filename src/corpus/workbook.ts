import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { buildCorpusWorkbook } from './helpers.ts'

/**
 * The last operations that needed no new assertion kind — plus the two that needed one.
 *
 * `autoFitColumnWidths` and `mailMerge` were simply missed: their results are a column
 * width and a filled-in sheet, both of which the existing assertions already read.
 * `unmerge` and `unmergeAll` needed a `merged` condition; `moveSheet` and
 * `setWorkbookProperties` needed a workbook-level subject, because neither is visible
 * from any sheet.
 *
 * Two operations are absent on purpose, and for the same reason: the corpus cannot
 * produce the input they need to be observable. `showFormulas` sets a flag in the sheet
 * view that ExcelJS does not surface on a round trip, and `freezeFormulas` only converts
 * formulas that carry a cached result — a workbook written by ExcelJS has none, so the
 * operation correctly skips them and warns. Asserting either would fail forever and read
 * as a broken operation.
 */
const order = (rows: Array<Array<string | number | null>>) => ({
  name: '订单',
  headers: ['产品', '数量'],
  rows,
})

const three = [['苹果', 10], ['香蕉', 5], ['梨', 7]]

export const workbookTasks: FileBenchmarkTask[] = [
  {
    id: 'workbook-autofit-columns',
    category: 'editing',
    name: '自动调整列宽',
    description: '产品名长短不一，让列宽自动适应内容。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workbook-autofit-columns', [order(three)]),
    operations: [{ op: 'autoFitColumnWidths', sheet: '订单', columns: ['A'], minWidth: 10, maxWidth: 30 }],
    checks: [
      { sheet: '订单', columnWidth: { A: 10 } },
      { id: '订单!A2', expect: '苹果' },
    ],
  },
  {
    id: 'workbook-merge-then-unmerge',
    category: 'editing',
    name: '合并表头后再拆开',
    description: '先合并两个表头格，发现不合适又拆回来。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workbook-merge-then-unmerge', [order(three)]),
    operations: [
      { op: 'mergeCells', range: '订单!A1:B1' },
      { op: 'unmerge', range: '订单!A1:B1' },
    ],
    checks: [
      { sheet: '订单', merged: [] },
      { id: '订单!A1', expect: '产品' },
    ],
  },
  {
    id: 'workbook-unmerge-all',
    category: 'editing',
    name: '拆掉表里所有合并格',
    description: '从别处拷来的表合并格很乱，一次全部拆掉。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workbook-unmerge-all', [order(three)]),
    operations: [
      { op: 'mergeCells', range: '订单!A1:B1' },
      { op: 'mergeCells', range: '订单!A2:B2' },
      { op: 'unmergeAll', sheet: '订单' },
    ],
    checks: [
      { sheet: '订单', merged: [] },
      { id: '订单!A1', expect: '产品' },
    ],
  },
  {
    id: 'workbook-move-sheet',
    category: 'editing',
    name: '调整工作表顺序',
    description: '把汇总表挪到最前面，打开文件先看到它。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workbook-move-sheet', [
      order(three),
      { name: '汇总', headers: ['说明'], rows: [['合计 22']] },
    ]),
    operations: [{ op: 'moveSheet', name: '汇总', position: 0 }],
    checks: [
      { workbook: true, sheetOrder: ['汇总', '订单'] },
      { id: '汇总!A2', expect: '合计 22' },
    ],
  },
  {
    id: 'workbook-document-properties',
    category: 'editing',
    name: '填写文档属性',
    description: '把标题和作者写上，方便归档检索。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workbook-document-properties', [order(three)]),
    operations: [
      { op: 'setWorkbookProperties', title: '2026 年订单台账', creator: '库管' },
    ],
    checks: [
      { workbook: true, properties: { title: '2026 年订单台账', creator: '库管' } },
      { id: '订单!A2', expect: '苹果' },
    ],
  },
  {
    id: 'workbook-mail-merge',
    category: 'editing',
    name: '按数据批量生成通知',
    description: '模板里留占位符，按数据表逐行生成通知文字。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'workbook-mail-merge', [
      { name: '模板', headers: ['内容'], rows: [['{产品} 还差 {数量} 件']] },
      { name: '数据', headers: ['产品', '数量'], rows: [['苹果', 10], ['香蕉', 5]] },
    ]),
    operations: [
      // A1 is the header; the template *body* is A2. The data range does start at the
      // header, because that row supplies the placeholder names.
      { op: 'mailMerge', template: '模板!A2', data: '数据!A1:B3', outputSheet: '通知' },
    ],
    checks: [
      // Each row of data becomes one line with the placeholders filled in — the merge
      // has to reach the data sheet, not just copy the template.
      // The merge writes from row 1 with no header of its own — the template supplies
      // the only header it has. Asserted at A1/A2 first, then wrongly moved to A2/A3
      // on the assumption that a header row must be there; it is not.
      { id: '通知!A1', expect: '苹果 还差 10 件' },
      { id: '通知!A2', expect: '香蕉 还差 5 件' },
    ],
  },
]
