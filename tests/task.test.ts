import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { access, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runExcelTask } from '../src/task.ts'

async function makeBrokenWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-task-'))
  const path = join(dir, 'book.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = '产品'
  sheet.getCell('B1').value = '收入'
  sheet.getCell('C1').value = '成本'
  sheet.getCell('D1').value = '毛利'
  sheet.getCell('D2').value = { formula: 'B2-C2' }
  sheet.getCell('D3').value = { formula: 'B3-C3' }
  sheet.getCell('D4').value = { formula: 'B4-C3' }
  sheet.getCell('D5').value = { formula: 'B5-C5' }
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('runExcelTask chains steps with per-step verify and auto-repair', async () => {
  const path = await makeBrokenWorkbook()
  const result = await runExcelTask(path, [
    {
      name: 'clean',
      operations: [
        { op: 'set', cells: { 'Sheet1!E1': '备注' } },
        { op: 'style', range: 'Sheet1!A1:E1', style: { bold: true } },
      ],
    },
    {
      name: 'summary',
      operations: [
        { op: 'set', cells: { 'Sheet1!F1': '合计' } },
      ],
    },
  ])
  assert.equal(result.steps.length, 2)
  assert.equal(result.steps[0]!.name, 'clean')
  assert.equal(result.steps[0]!.validation!.before, 1)
  assert.equal(result.steps[0]!.validation!.fixed, 1)
  assert.equal(result.steps[0]!.validation!.after, 0)
  assert.equal(result.steps[1]!.validation!.before, 0)
  assert.equal(result.finalAnomalies, 0)
  assert.match(result.outputPath, /\.task\.xlsx$/)
  await access(result.outputPath)
})

test('runExcelTask skips validation when verify is false', async () => {
  const path = await makeBrokenWorkbook()
  const result = await runExcelTask(path, [{ name: 'no-verify', operations: [{ op: 'set', cells: { 'Sheet1!E1': 'x' } }], verify: false }])
  assert.equal(result.steps[0]!.validation, undefined)
})
