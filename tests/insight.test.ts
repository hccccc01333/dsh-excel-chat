import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildWorkbookInsight } from '../src/insight.ts'

async function makeWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-insight-'))
  const path = join(dir, 'sales.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '产品'
  sheet.getCell('B1').value = '数量'
  sheet.getCell('C1').value = '金额'
  const values = [1, 2, 3, 4, 5, 100]
  for (let row = 2; row <= 7; row++) {
    sheet.getCell(`A${row}`).value = row === 4 ? null : row === 5 ? '  橘子 ' : `产品${row}`
    sheet.getCell(`B${row}`).value = values[row - 2]
    sheet.getCell(`C${row}`).value = { formula: `B${row}*10` }
  }
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('buildWorkbookInsight reports missing, outlier, whitespace, and formula findings', async () => {
  const path = await makeWorkbook()
  const insight = await buildWorkbookInsight(path)
  assert.ok(insight.summary.includes('工作表'))
  assert.equal(insight.sheets.length, 1)
  const findings = insight.sheets[0]!.findings
  const categories = findings.map((finding) => finding.category)
  assert.ok(categories.includes('missing'), 'missing finding expected')
  assert.ok(categories.includes('outlier'), 'outlier finding expected')
  assert.ok(categories.includes('whitespace'), 'whitespace finding expected')
  assert.ok(categories.includes('formula'), 'formula finding expected')
  const suggestions = insight.suggestions.join('\n')
  assert.ok(suggestions.includes('fillMissing'))
  assert.ok(suggestions.includes('excel_autofix'))
})

test('buildWorkbookInsight reports a clean table without warnings', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-insight-clean-'))
  const path = join(dir, 'clean.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = '名称'
  sheet.getCell('B1').value = '数量'
  sheet.getCell('A2').value = '苹果'
  sheet.getCell('B2').value = 10
  sheet.getCell('A3').value = '香蕉'
  sheet.getCell('B3').value = 20
  await writeFile(path, await workbook.xlsx.writeBuffer())
  const insight = await buildWorkbookInsight(path)
  const warnings = insight.sheets.flatMap((entry) => entry.findings).filter((finding) => finding.severity !== 'info')
  assert.equal(warnings.length, 0)
  assert.ok(insight.suggestions.join('\n').includes('未发现明显数据问题'))
})
