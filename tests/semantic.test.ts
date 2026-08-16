import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import ExcelJS from 'exceljs'
import { buildWorkbookSemanticProfile } from '../src/semantic.ts'

async function makeWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-semantic-'))
  const path = join(dir, 'sales.xlsx')
  const workbook = new ExcelJS.Workbook()
  const orders = workbook.addWorksheet('订单')
  for (const [col, header] of ['日期', '区域', '产品', '数量', '单价', '金额'].entries()) {
    orders.getCell(`${String.fromCharCode(65 + col)}1`).value = header
  }
  orders.getCell('A2').value = '2026-01-05'
  orders.getCell('B2').value = '华东'
  orders.getCell('C2').value = 'A'
  orders.getCell('D2').value = 10
  orders.getCell('E2').value = 100
  orders.getCell('F2').value = { formula: 'D2*E2' }
  const prices = workbook.addWorksheet('产品价目')
  prices.getCell('A1').value = '产品'
  prices.getCell('B1').value = '名称'
  prices.getCell('A2').value = 'A'
  prices.getCell('B2').value = '台式机'
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('buildWorkbookSemanticProfile classifies roles, grain, derived metrics, and join keys', async () => {
  const path = await makeWorkbook()
  const profile = await buildWorkbookSemanticProfile(path)
  const orders = profile.sheets.find((sheet) => sheet.sheet === '订单')
  assert.ok(orders)
  const roles = Object.fromEntries(orders.columns.map((column) => [column.header, column.role]))
  assert.equal(roles['日期'], 'time')
  assert.equal(roles['区域'], 'dimension')
  assert.equal(roles['产品'], 'dimension')
  assert.equal(roles['数量'], 'measure')
  assert.equal(roles['单价'], 'measure')
  assert.equal(roles['金额'], 'measure')
  assert.match(orders.grain, /日期/)
  assert.match(orders.grain, /区域/)
  assert.ok(orders.derivedMetrics.some((metric) => metric.includes('金额') && metric.includes('D2*E2')))
  assert.ok(profile.joinKeys.some((key) => key.left === '订单' && key.right === '产品价目' && key.key === '产品'))
  assert.match(profile.summary, /指标=[^；]*金额/)
  assert.match(profile.summary, /可关联：订单.产品 ↔ 产品价目.产品/)
})
