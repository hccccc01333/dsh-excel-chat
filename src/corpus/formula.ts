import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { buildCorpusWorkbook } from './helpers.ts'

const margin = (rows: Array<Array<string | number | null>>) => ({
  name: '订单',
  headers: ['产品', '收入', '成本', '毛利'],
  rows,
})

export const formulaTasks: FileBenchmarkTask[] = [
  {
    id: 'formula-repair-offset',
    category: 'formula',
    name: '修复引用偏移',
    description: '毛利列一处公式引用错行，验证后自动修复。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'formula-repair-offset', [margin([
      ['A', 100, 60, '=B2-C2'],
      ['B', 200, 120, '=B3-C3'],
      ['C', 300, 150, '=B4-C3'],
      ['D', 400, 200, '=B5-C5'],
    ])]),
    operations: [],
    evaluateAfterAutofix: true,
    checks: [
      { id: '订单!D4', expect: '=B4-C4' },
      { id: '订单!D5', expect: '=B5-C5' },
    ],
  },
  {
    id: 'formula-repair-range-tail',
    category: 'formula',
    name: '修复范围尾端',
    description: 'SUM 范围尾端引用错行，自动修复。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'formula-repair-range-tail', [margin([
      ['A', 100, 60, '=SUM(B2:C2)'],
      ['B', 200, 120, '=SUM(B3:C3)'],
      ['C', 300, 150, '=SUM(B4:C3)'],
      ['D', 400, 200, '=SUM(B5:C5)'],
    ])]),
    operations: [],
    evaluateAfterAutofix: true,
    checks: [{ id: '订单!D4', expect: '=SUM(B4:C4)' }],
  },
  {
    id: 'formula-repair-empty-gap',
    category: 'formula',
    name: '补全空行公式',
    description: '毛利列缺一行公式，自动按相邻公式补全。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'formula-repair-empty-gap', [margin([
      ['A', 100, 60, '=B2-C2'],
      ['B', 200, 120, '=B3-C3'],
      ['C', 300, 150, null],
      ['D', 400, 200, '=B5-C5'],
    ])]),
    operations: [],
    evaluateAfterAutofix: true,
    checks: [{ id: '订单!D4', startsWith: '=' }],
  },
  {
    id: 'formula-vlookup',
    category: 'formula',
    name: 'VLOOKUP 补列',
    description: '从价目表按编码补产品名称。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'formula-vlookup', [
      {
        name: '订单',
        headers: ['编码', '数量', '产品名称'],
        rows: [['P01', 10], ['P02', 5]],
      },
      { name: '价目表', headers: ['编码', '名称'], rows: [['P01', '苹果'], ['P02', '香蕉']] },
    ]),
    operations: [{
      op: 'set',
      cells: {
        '订单!C2': '=VLOOKUP(A2,价目表!$A$1:$B$3,2,FALSE)',
        '订单!C3': '=VLOOKUP(A3,价目表!$A$1:$B$3,2,FALSE)',
      },
    }],
    checks: [{ id: '订单!C2', startsWith: '=VLOOKUP(A2,价目表' }],
  },
  {
    id: 'formula-sumifs',
    category: 'formula',
    name: 'SUMIFS 条件汇总',
    description: '写 SUMIFS 汇总华东金额。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'formula-sumifs', [{
      name: '订单',
      headers: ['区域', '金额'],
      rows: [['华东', 100], ['华北', 200], ['华东', 300]],
    }]),
    operations: [{
      op: 'set',
      cells: { '订单!D2': '=SUMIFS(B2:B4,A2:A4,"华东")' },
    }],
    checks: [{ id: '订单!D2', startsWith: '=SUMIFS(' }],
  },
  {
    id: 'formula-if',
    category: 'formula',
    name: 'IF 达标判断',
    description: '写 IF 判断金额是否达标。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'formula-if', [{
      name: '订单',
      headers: ['产品', '金额', '状态'],
      rows: [['A', 100], ['B', 300]],
    }]),
    operations: [{
      op: 'set',
      cells: {
        '订单!C2': '=IF(B2>200,"达标","未达标")',
        '订单!C3': '=IF(B3>200,"达标","未达标")',
      },
    }],
    checks: [{ id: '订单!C2', startsWith: '=IF(' }],
  },
]
