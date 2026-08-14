import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { applyOperationsToWorkbook, type ExcelOperation } from '../src/operations.ts'
import { readWorkbookCells } from '../src/workbook.ts'

async function makeSalesWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-scenarios-'))
  const path = join(dir, '销售台账.xlsx')
  const workbook = new ExcelJS.Workbook()
  const orders = workbook.addWorksheet('订单')
  orders.getCell('A1').value = '日期'
  orders.getCell('B1').value = '区域'
  orders.getCell('C1').value = '产品'
  orders.getCell('D1').value = '数量'
  orders.getCell('E1').value = '单价'
  orders.getCell('F1').value = '金额'
  const rows: Array<Array<string | number>> = [
    ['2026-01-05', '华东', 'A', 10, 100],
    ['2026-01-08', '华北', 'B', 5, 200],
    ['2026-01-12', '华东', 'B', 8, 200],
    ['2026-01-15', '华南', 'A', 4, 100],
    ['2026-01-19', '华北', 'A', 12, 100],
    ['2026-01-22', '华东', 'C', 6, 300],
  ]
  rows.forEach((row, index) => {
    const r = index + 2
    orders.getCell(`A${r}`).value = row[0] as string
    orders.getCell(`B${r}`).value = row[1] as string
    orders.getCell(`C${r}`).value = row[2] as string
    orders.getCell(`D${r}`).value = row[3] as number
    orders.getCell(`E${r}`).value = row[4] as number
    orders.getCell(`F${r}`).value = { formula: `D${r}*E${r}` }
  })
  const prices = workbook.addWorksheet('产品价目')
  prices.getCell('A1').value = '产品'
  prices.getCell('B1').value = '名称'
  prices.getCell('A2').value = 'A'
  prices.getCell('B2').value = '台式机'
  prices.getCell('A3').value = 'B'
  prices.getCell('B3').value = '笔记本'
  prices.getCell('A4').value = 'C'
  prices.getCell('B4').value = '显示器'
  const notice = workbook.addWorksheet('通知模板')
  notice.getCell('A1').value = '{区域}'
  notice.getCell('B1').value = '数量 {数量}'
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

async function apply(path: string, operations: ExcelOperation[], outName: string): Promise<Record<string, string>> {
  const outPath = join(join(path, '..'), outName)
  await applyOperationsToWorkbook(path, operations, outPath)
  return readWorkbookCells(await readFile(outPath))
}

test('场景1：按区域生成动态透视报表（SUMIFS 联动）', async () => {
  const path = await makeSalesWorkbook()
  const cells = await apply(path, [{
    op: 'aggregateReport',
    source: '订单!A1:F7',
    groupColumn: 'B',
    metrics: [
      { column: 'F', function: 'sum' },
      { column: 'D', function: 'count' },
    ],
  }], '透视.xlsx')
  assert.equal(cells['订单-汇总!A1'], '区域')
  assert.equal(cells['订单-汇总!B1'], '金额 合计')
  assert.equal(cells['订单-汇总!A2'], '华东')
  assert.equal(cells['订单-汇总!B2'], '=SUMIFS(订单!$F$2:$F$7,订单!$B$2:$B$7,A2)')
  assert.equal(cells['订单-汇总!C2'], '=COUNTIFS(订单!$D$2:$D$7,订单!$B$2:$B$7,A2)')
  assert.equal(cells['订单-汇总!A5'], '总计')
})

test('场景2：VLOOKUP 从价目表补产品名称', async () => {
  const path = await makeSalesWorkbook()
  const cells = await apply(path, [{
    op: 'set',
    cells: {
      '订单!G1': '产品名称',
      '订单!G2': '=VLOOKUP(C2,产品价目!$A$1:$B$4,2,0)',
      '订单!G3': '=VLOOKUP(C3,产品价目!$A$1:$B$4,2,0)',
      '订单!G4': '=VLOOKUP(C4,产品价目!$A$1:$B$4,2,0)',
      '订单!G5': '=VLOOKUP(C5,产品价目!$A$1:$B$4,2,0)',
      '订单!G6': '=VLOOKUP(C6,产品价目!$A$1:$B$4,2,0)',
      '订单!G7': '=VLOOKUP(C7,产品价目!$A$1:$B$4,2,0)',
    },
  }], 'vlookup.xlsx')
  assert.equal(cells['订单!G1'], '产品名称')
  assert.equal(cells['订单!G2'], '=VLOOKUP(C2,产品价目!$A$1:$B$4,2,0)')
  assert.equal(cells['订单!G7'], '=VLOOKUP(C7,产品价目!$A$1:$B$4,2,0)')
})

test('场景3：先排序再分类汇总，得到分区域小计与总计', async () => {
  const path = await makeSalesWorkbook()
  const cells = await apply(path, [
    { op: 'sortRange', range: '订单!A1:F7', keys: [{ column: 'B' }], headerRows: 1 },
    {
      op: 'subtotal',
      sheet: '订单',
      range: '订单!A1:F7',
      groupColumn: 'B',
      summaryColumns: [{ column: 'F', function: 'sum' }],
    },
  ], '汇总.xlsx')
  const hasSubtotal = Object.values(cells).includes('华东 汇总')
  const hasTotal = Object.values(cells).includes('总计')
  const hasSubtotalFormula = Object.values(cells).some((content) => content.startsWith('=SUBTOTAL(9,F'))
  assert.ok(hasSubtotal, '缺少分区域汇总行')
  assert.ok(hasTotal, '缺少总计行')
  assert.ok(hasSubtotalFormula, '缺少 SUBTOTAL 公式')
})

test('场景4：邮件合并批量通知 + 工作表保护 + 命名区域', async () => {
  const path = await makeSalesWorkbook()
  const outPath = join(join(path, '..'), '合并.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'mailMerge', template: '通知模板!A1:B1', data: '订单!A1:D7', outputSheet: '发货通知' },
    { op: 'protectSheet', sheet: '订单', password: 'pw', options: { sort: true } },
    { op: 'definedName', name: 'SalesData', ref: '订单!$A$1:$F$7' },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const notice = workbook.getWorksheet('发货通知')!
  assert.equal(notice.getCell('A1').value, '华东')
  assert.equal(notice.getCell('B1').value, '数量 10')
  assert.equal(notice.getCell('A6').value, '华东')
  assert.equal(notice.getCell('B6').value, '数量 6')
  const names = workbook.definedNames.model as Array<{ name: string; ranges: string[] }>
  assert.ok(names.some((entry) => entry.name === 'SalesData'))
})
