import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import ExcelJS from 'exceljs'
import { writeWorkbookHealthReport } from '../src/health-report.ts'
import { validate } from '../src/validator.ts'
import { readWorkbookCells } from '../src/workbook.ts'

async function makeBrokenWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-health-'))
  const path = join(dir, 'sales.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '区域'
  sheet.getCell('B1').value = '数量'
  sheet.getCell('C1').value = '单价'
  sheet.getCell('D1').value = '金额'
  sheet.getCell('A2').value = '华东'
  sheet.getCell('B2').value = 2
  sheet.getCell('C2').value = 10
  sheet.getCell('D2').value = { formula: 'B2*C2' }
  sheet.getCell('A3').value = '华北'
  sheet.getCell('B3').value = 3
  sheet.getCell('C3').value = 20
  sheet.getCell('D3').value = { formula: 'B3*C3' }
  sheet.getCell('A4').value = '华南'
  sheet.getCell('B4').value = 4
  sheet.getCell('C4').value = 30
  sheet.getCell('D4').value = { formula: 'B4*C3' }
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('writeWorkbookHealthReport embeds a hidden report sheet with score and anomalies', async () => {
  const path = await makeBrokenWorkbook()
  const result = await writeWorkbookHealthReport(path)
  assert.equal(result.reportSheet, '_dsh_体检报告')
  assert.ok(result.healthScore < 100)
  assert.ok(result.anomalyCount >= 1)
  assert.match(result.summary, /_dsh_体检报告/)

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(path)
  const report = workbook.getWorksheet('_dsh_体检报告')
  assert.ok(report)
  assert.equal(report.state, 'hidden')
  assert.match(String(report.getCell('A6').value), /健康分：/)
})

test('validator ignores _dsh_ internal sheets', async () => {
  const path = await makeBrokenWorkbook()
  await writeWorkbookHealthReport(path)
  const cells = await readWorkbookCells(await readFile(path))
  const result = validate(cells)
  assert.ok(Object.keys(cells).some((id) => id.toUpperCase().includes('_DSH_')))
  assert.ok(result.anomalies.every((anomaly) => !anomaly.cell.toUpperCase().includes('_DSH_')))
  assert.ok(result.formulaCount >= 1)
})
