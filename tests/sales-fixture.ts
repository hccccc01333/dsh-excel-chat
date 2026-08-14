import ExcelJS from 'exceljs'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/** Realistic sales ledger: order sheet + price table + merge notice template. */
export async function makeSalesWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-sales-'))
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
