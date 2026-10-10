import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { buildCorpusWorkbook } from './helpers.ts'

/**
 * Inserting and deleting rows and columns.
 *
 * Excel moves every formula that points past the edit, and nothing else in the
 * workbook does — so this is the operation family where a wrong answer looks right.
 * None of it was in the corpus: of the 77 operations the schema allows, these five
 * were among the fifty never used by a task, and the reference-shifting helpers in
 * `op/structure.ts` had no end-to-end coverage at all.
 *
 * The expectations here were read off the implementation before they were written
 * down, then kept because they are the behaviour a user would want: a formula that
 * pointed at a cell still points at the same cell after the edit.
 */
const order = (rows: Array<Array<string | number | null>>) => ({
  name: '订单',
  headers: ['产品', '数量'],
  rows,
})

export const structureTasks: FileBenchmarkTask[] = [
  {
    id: 'structure-insert-rows',
    category: 'editing',
    name: '插入一行',
    description: '在订单表第 2 行前插入一个空行，原有数据下移。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'structure-insert-rows', [order([
      ['苹果', 10], ['香蕉', 5],
    ])]),
    operations: [{ op: 'insertRows', sheet: '订单', row: 2, count: 1 }],
    checks: [
      { id: '订单!A2', expect: null },
      { id: '订单!A3', expect: '苹果' },
      // Conjunction: the row moved *and* its number came with it. An insert that
      // shifted the labels but not the values would pass either half alone.
      { id: '订单!B3', expect: '10' },
      { id: '订单!A4', expect: '香蕉' },
    ],
  },
  {
    id: 'structure-insert-rows-shifts-formula',
    category: 'editing',
    name: '插入行时公式跟随位移',
    description: '第 2 行插入空行后，原本引用 B2 的公式应当改为引用 B3。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'structure-insert-rows-shifts-formula', [{
      name: '订单',
      headers: ['产品', '数量', '小计'],
      rows: [['苹果', 10, '=B2*2'], ['香蕉', 5, null]],
    }]),
    operations: [{ op: 'insertRows', sheet: '订单', row: 2, count: 1 }],
    checks: [
      // The whole point: the formula follows the data it referenced.
      { id: '订单!C3', startsWith: '=B3*2' },
      { id: '订单!B3', expect: '10' },
    ],
  },
  {
    id: 'structure-delete-rows',
    category: 'editing',
    name: '删除一行',
    description: '删除订单表第 2 行，后面的行上移。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'structure-delete-rows', [order([
      ['苹果', 10], ['香蕉', 5], ['梨', 3],
    ])]),
    operations: [{ op: 'deleteRows', sheet: '订单', row: 2, count: 1 }],
    checks: [
      { id: '订单!A2', expect: '香蕉' },
      { id: '订单!B2', expect: '5' },
      { id: '订单!A3', expect: '梨' },
      // The deleted row is gone, not blanked in place.
      { id: '订单!A4', expect: null },
    ],
  },
  {
    id: 'structure-insert-columns',
    category: 'editing',
    name: '插入一列',
    description: '在订单表 A 列前插入一个空列，原有列右移。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'structure-insert-columns', [order([
      ['苹果', 10], ['香蕉', 5],
    ])]),
    operations: [{ op: 'insertColumns', sheet: '订单', column: 'A', count: 1 }],
    checks: [
      { id: '订单!A1', expect: null },
      { id: '订单!B1', expect: '产品' },
      { id: '订单!C1', expect: '数量' },
      { id: '订单!B2', expect: '苹果' },
    ],
  },
  {
    id: 'structure-delete-columns-shifts-formula',
    category: 'editing',
    name: '删除列时公式跟随位移',
    description: '删除 A 列后，原本引用 B2 的公式应当改为引用 A2。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'structure-delete-columns-shifts-formula', [{
      name: '订单',
      headers: ['产品', '数量', '小计'],
      rows: [['苹果', 10, '=B2*2'], ['香蕉', 5, '=B3*2']],
    }]),
    operations: [{ op: 'deleteColumns', sheet: '订单', column: 'A', count: 1 }],
    checks: [
      { id: '订单!A1', expect: '数量' },
      { id: '订单!A2', expect: '10' },
      // Column references move too, not just row references.
      { id: '订单!B2', startsWith: '=A2*2' },
    ],
  },
]
