import { readFile } from 'node:fs/promises'
import ExcelJS from 'exceljs'
import { t } from './i18n.ts'
import { validate } from './validator.ts'
import { readWorkbookCells, stripPivotTableParts } from './workbook.ts'
import { writeWorkbookSafely } from './safe-write.ts'

/**
 * Sheet name is an identifier, not prose: it is written into the file, exported
 * here, and skipped by `validator` through its `_dsh_` prefix. It therefore stays
 * the same in every language — only the report's *content* is translated.
 */
export const HEALTH_REPORT_SHEET = '_dsh_体检报告'

export interface HealthReportResult {
  path: string
  healthScore: number
  formulaCount: number
  anomalyCount: number
  reportSheet: string
  summary: string
}

/**
 * Write a formula health report INTO the workbook itself: a hidden
 * `_dsh_体检报告` sheet with score, counts, and per-anomaly rows. The file
 * carries its own audit trail, independent of chat history or external logs.
 * Validators skip `_dsh_` sheets so the report never flags itself.
 */
export async function writeWorkbookHealthReport(path: string, outPath?: string): Promise<HealthReportResult> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(stripPivotTableParts(await readFile(path)) as any)
  const cells = await readWorkbookCells(await readFile(path))
  const result = validate(cells)
  const anomalyCount = result.anomalies.length
  const formulaCount = result.formulaCount
  const healthScore = Math.max(0, 100 - anomalyCount * 10)

  const existing = workbook.getWorksheet(HEALTH_REPORT_SHEET)
  if (existing) workbook.removeWorksheet(existing.id)
  const report = workbook.addWorksheet(HEALTH_REPORT_SHEET, { state: 'hidden' })
  const bold = { bold: true } as Partial<ExcelJS.Font>
  const set = (cell: string, value: ExcelJS.CellValue, font?: Partial<ExcelJS.Font>): void => {
    report.getCell(cell).value = value
    if (font) report.getCell(cell).font = font
  }
  set('A1', t('公式健康报告'), bold)
  set('A2', t('生成时间：{time}', { time: new Date().toISOString() }))
  set('A3', t('文件：{path}', { path }))
  set('A4', t('公式数：{count}', { count: formulaCount }))
  set('A5', t('异常数：{count}', { count: anomalyCount }))
  set('A6', t('健康分：{score}', { score: healthScore }))
  set('A8', t('单元格'), bold)
  set('B8', t('类型'), bold)
  set('C8', t('说明'), bold)
  for (const [index, anomaly] of result.anomalies.slice(0, 200).entries()) {
    const row = 9 + index
    set(`A${row}`, anomaly.cell)
    set(`B${row}`, anomaly.kind)
    set(`C${row}`, anomaly.message)
  }
  report.getColumn(1).width = 24
  report.getColumn(2).width = 18
  report.getColumn(3).width = 80

  const target = outPath ?? path
  await writeWorkbookSafely(target, new Uint8Array(await workbook.xlsx.writeBuffer()))
  const summary = t('健康分 {score}：{formulas} 个公式，{anomalies} 个异常，报告已写入 {path} 的「{sheet}」表', {
    score: healthScore,
    formulas: formulaCount,
    anomalies: anomalyCount,
    path: target,
    sheet: HEALTH_REPORT_SHEET,
  })
  return { path: target, healthScore, formulaCount, anomalyCount, reportSheet: HEALTH_REPORT_SHEET, summary }
}
