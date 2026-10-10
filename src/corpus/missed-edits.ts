import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { buildCorpusWorkbook } from './helpers.ts'

/**
 * Edits the corpus had never asked for.
 *
 * Every one of these is in the schema and reached by a user asking in plain language —
 * "去个重"、"加一列排名"、"转置一下"、"把这块清空"、"删掉空列"、"格式刷一下"、"把这个 CSV
 * 导进来" — and none of them appeared in a task.
 *
 *  and  are deliberately absent: they name a file outside the
 * workbook, and a corpus task's operations are static while the fixture directory is
 * only known at build time. They are covered by unit tests instead. Expectations were read off the
 * implementation before being written down.
 */
const order = (rows: Array<Array<string | number | null>>) => ({
  name: '订单',
  headers: ['产品', '数量'],
  rows,
})

export const missedEditTasks: FileBenchmarkTask[] = [
  {
    id: 'edit-unique-values-first-seen',
    category: 'editing',
    name: '去重并保留首次出现顺序',
    description: '把产品列的不重复值写到旁边，保持第一次出现的顺序。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-unique-values-first-seen', [{
      name: '订单',
      headers: ['产品'],
      rows: [['苹果'], ['苹果'], ['香蕉'], ['梨'], ['香蕉']],
    }]),
    operations: [{ op: 'uniqueValues', source: '订单!A1:A6', target: '订单!C1', includeHeader: true }],
    checks: [
      { id: '订单!C1', expect: '产品' },
      { id: '订单!C2', expect: '苹果' },
      // 香蕉 before 梨: first-seen order, not sorted order. A sort would swap them.
      { id: '订单!C3', expect: '香蕉' },
      { id: '订单!C4', expect: '梨' },
      { id: '订单!C5', expect: null },
    ],
  },
  {
    id: 'edit-rank-column-live-formulas',
    category: 'editing',
    name: '加一列排名',
    description: '按数量给每行加排名公式。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-rank-column-live-formulas', [order([
      ['苹果', 10], ['香蕉', 5],
    ])]),
    operations: [{ op: 'rankColumn', range: '订单!A1:B3', metricColumn: 'B', outputColumn: 'C' }],
    checks: [
      { id: '订单!C2', startsWith: '=RANK(' },
      { id: '订单!C3', startsWith: '=RANK(' },
      // The metric is untouched.
      { id: '订单!B2', expect: '10' },
    ],
  },
  {
    id: 'edit-transpose-block',
    category: 'editing',
    name: '转置一块数据',
    description: '把 2 列 3 行的表转置到旁边。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-transpose-block', [order([
      ['苹果', 10], ['香蕉', 5],
    ])]),
    operations: [{ op: 'transpose', source: '订单!A1:B3', target: '订单!D1' }],
    checks: [
      { id: '订单!D1', expect: '产品' },
      { id: '订单!E1', expect: '苹果' },
      { id: '订单!D2', expect: '数量' },
      { id: '订单!E2', expect: '10' },
      // The source stays put — transpose copies, it does not move.
      { id: '订单!A1', expect: '产品' },
    ],
  },
  {
    id: 'edit-clear-range-keeps-neighbours',
    category: 'editing',
    name: '清空数值列但保留产品列',
    description: '只清空数量列的内容。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-clear-range-keeps-neighbours', [order([
      ['苹果', 10], ['香蕉', 5],
    ])]),
    operations: [{ op: 'clearRange', range: '订单!B2:B3', mode: 'contents' }],
    checks: [
      { id: '订单!B2', expect: null },
      { id: '订单!B3', expect: null },
      // The column beside it is untouched: a clear that overshot would pass alone.
      { id: '订单!A2', expect: '苹果' },
      { id: '订单!A3', expect: '香蕉' },
    ],
  },
  {
    id: 'edit-remove-empty-columns',
    category: 'editing',
    name: '删掉整列为空的列',
    description: '去掉中间那一列全空的数据。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'edit-remove-empty-columns', [{
      name: '订单',
      // The middle column is empty *including its header*: a header would be content,
      // and the operation is right to keep a column that has any.
      headers: ['产品', '', '数量'],
      rows: [['苹果', null, 10], ['香蕉', null, 5]],
    }]),
    operations: [{ op: 'removeEmptyColumns', range: '订单!A1:C3' }],
    checks: [
      { id: '订单!A1', expect: '产品' },
      // 数量 moves left into the removed column's place.
      { id: '订单!B1', expect: '数量' },
      { id: '订单!B2', expect: '10' },
      { id: '订单!C1', expect: null },
    ],
  },
  {
    id: 'format-copy-style-to-range',
    category: 'editing',
    name: '格式刷到整行表头',
    description: '把 A1 的格式复制到 B1。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'format-copy-style-to-range', [{
      name: '订单',
      headers: ['产品', '数量'],
      rows: [['苹果', 10]],
      styles: { A1: { bold: true, fill: 'FFFF00' } },
    }]),
    operations: [{ op: 'copyStyle', source: '订单!A1', target: '订单!A1:B1' }],
    checks: [
      // Style only: B1 has no value of its own, so this asserts the copy happened
      // rather than something the sheet already had.
      { id: '订单!B1', bold: true, fill: 'FFFF00' },
      { id: '订单!A1', bold: true },
    ],
  },
]
