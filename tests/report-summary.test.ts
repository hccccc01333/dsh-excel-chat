import assert from 'node:assert/strict'
import { test } from 'node:test'
import ExcelJS from 'exceljs'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { operateWorkbookFile } from '../src/operations.ts'

/**
 * `report` inserts subtotal rows into the source sheet *before* building its
 * summary sheet, so the block it hands to `aggregateReport` contains derived
 * rows as well as data. A subtotal row's label ("华东 汇总", "总计") is a
 * different string from every real group key, so enumerating the group column
 * naively turned each one into a phantom group — and the grand-total row then
 * summed the real groups twice over (a total three times the truth).
 *
 * These assertions pin the structure, because that is where the defect lived;
 * Excel evaluates the formulas to the correct numbers once the shape is right.
 */
async function buildReport() {
  const dir = await mkdtemp(join(tmpdir(), 'report-summary-'))
  const path = join(dir, 'sales.xlsx')
  const wb = new ExcelJS.Workbook()
  const sheet = wb.addWorksheet('订单')
  sheet.addRow(['订单号', '区域', '渠道', '数量', '单价', '金额'])
  const regions = ['华东', '华北', '华南', '西南']
  for (let i = 0; i < 24; i++) {
    sheet.addRow([`A-${1000 + i}`, regions[i % 4], '线上', 5 + ((i * 7) % 40), 20 + ((i * 13) % 180), { formula: `D${i + 2}*E${i + 2}` }])
  }
  await wb.xlsx.writeFile(path)

  const out = await operateWorkbookFile(path, [
    { op: 'report', source: '订单!A1:F25', groupColumn: 'B', metrics: [{ column: 'F', function: 'sum' }], outputSheet: '区域汇总' },
  ], join(dir, 'report.xlsx'))

  const summary = new ExcelJS.Workbook()
  await summary.xlsx.load(await import('node:fs/promises').then((fs) => fs.readFile(out.outputPath)))
  return summary.getWorksheet('区域汇总')!
}

test('the report summary lists each group once, without phantom subtotal rows', async () => {
  const sheet = await buildReport()
  const labels: string[] = []
  for (let row = 2; row <= sheet.rowCount; row++) labels.push(String(sheet.getCell(row, 1).value))

  assert.deepEqual(labels, ['华东', '华北', '华南', '西南', '总计'])
  for (const label of labels) {
    assert.doesNotMatch(label, /汇总/, `"${label}" is a subtotal row leaking into the summary`)
  }
})

test('the report grand total sums the group rows, not the subtotal rows', async () => {
  const sheet = await buildReport()
  const totalRow = sheet.rowCount
  assert.equal(String(sheet.getCell(totalRow, 1).value), '总计')

  const formula = (sheet.getCell(totalRow, 2).value as { formula?: string }).formula
  // Group rows are 2..totalRow-1; a wider range would re-count the subtotals.
  assert.equal(formula, `SUM(B2:B${totalRow - 1})`)

  // Every row above the total must be a real SUMIFS group, never a SUM of SUMS.
  for (let row = 2; row < totalRow; row++) {
    const cell = sheet.getCell(row, 2).value as { formula?: string }
    assert.match(String(cell.formula), /^SUMIFS\(/, `row ${row} is not a group aggregate`)
  }
})

test('a standalone aggregateReport also skips subtotal rows in its source', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'report-standalone-'))
  const path = join(dir, 'sheet.xlsx')
  const wb = new ExcelJS.Workbook()
  const sheet = wb.addWorksheet('数据')
  sheet.addRow(['区域', '金额'])
  sheet.addRow(['华东', 10])
  sheet.addRow(['华东', 20])
  sheet.addRow(['华东 汇总', { formula: 'SUBTOTAL(9,B2:B3)' }])
  sheet.addRow(['华北', 5])
  sheet.addRow(['总计', { formula: 'SUBTOTAL(9,B2:B5)' }])
  await wb.xlsx.writeFile(path)

  const out = await operateWorkbookFile(path, [
    { op: 'aggregateReport', source: '数据!A1:B6', groupColumn: 'A', metrics: [{ column: 'B', function: 'sum' }], outputSheet: '汇总' },
  ], join(dir, 'out.xlsx'))

  const summary = new ExcelJS.Workbook()
  await summary.xlsx.load(await import('node:fs/promises').then((fs) => fs.readFile(out.outputPath)))
  const outSheet = summary.getWorksheet('汇总')!
  const labels: string[] = []
  for (let row = 2; row <= outSheet.rowCount; row++) labels.push(String(outSheet.getCell(row, 1).value))
  assert.deepEqual(labels, ['华东', '华北', '总计'])
})
