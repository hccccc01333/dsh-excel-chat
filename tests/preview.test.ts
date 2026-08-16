import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { access, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildWorkbookPreview } from '../src/preview.ts'

async function makeWorkbook(rows: number): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-preview-'))
  const path = join(dir, 'sales.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '产品'
  sheet.getCell('B1').value = '数量'
  sheet.getCell('C1').value = '金额'
  for (let row = 2; row <= rows + 1; row++) {
    sheet.getCell(`A${row}`).value = `产品${row}`
    sheet.getCell(`B${row}`).value = row
    sheet.getCell(`C${row}`).value = { formula: `B${row}*10` }
  }
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

test('buildWorkbookPreview returns a markdown table and writes an HTML preview', async () => {
  const path = await makeWorkbook(3)
  const result = await buildWorkbookPreview(path, { sheet: '订单' })
  assert.ok(result.markdown.includes('产品'))
  assert.ok(result.markdown.includes('=B2*10'))
  assert.ok(result.summary.includes('订单'))
  assert.equal(result.sheets.length, 1)
  assert.ok(result.sheets[0]!.cells.length > 0)
  assert.match(result.previewPath, /\.preview\.html$/)
  await access(result.previewPath)
  const html = await readFile(result.previewPath, 'utf8')
  assert.ok(html.includes('<table>'))
  assert.ok(html.includes('产品'))
})

test('buildWorkbookPreview honors maxRows and escapes pipes', async () => {
  const path = await makeWorkbook(30)
  const result = await buildWorkbookPreview(path, { range: 'A1:C31', maxRows: 5 })
  const dataLines = result.markdown.split('\n').filter((line) => line.startsWith('|') && line.includes('产品'))
  assert.equal(dataLines.length, 5)
  const dir = await mkdtemp(join(tmpdir(), 'vera-preview-pipe-'))
  const pipePath = join(dir, 'pipe.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = '备注'
  sheet.getCell('A2').value = 'a|b'
  await writeFile(pipePath, await workbook.xlsx.writeBuffer())
  const escaped = await buildWorkbookPreview(pipePath)
  assert.ok(escaped.markdown.includes('a\\|b'))
})

test('buildWorkbookPreview result is lossless JSON for the dsh harness', async () => {
  const path = await makeWorkbook(3)
  const result = await buildWorkbookPreview(path, { sheet: '订单' })
  const seen = new Set<object>()
  const walk = (value: unknown): void => {
    if (value === null || typeof value !== 'object') return
    assert.ok(!seen.has(value), 'circular reference')
    seen.add(value)
    for (const [key, child] of Object.entries(value)) {
      assert.notEqual(child, undefined, `field ${key} must not be undefined`)
      walk(child)
    }
  }
  walk(result)
  for (const sheet of result.sheets) {
    for (const cell of sheet.cells) {
      if (cell.formula === undefined) continue
      assert.match(cell.formula, /^=/, `formula ${cell.formula} must start with =`)
    }
  }
})
