import { profileWorkbook, type SheetProfile, type WorkbookProfile } from './profile.ts'

export type InsightSeverity = 'info' | 'warn' | 'alert'

export interface InsightFinding {
  severity: InsightSeverity
  category: string
  message: string
}

export interface SheetInsight {
  sheet: string
  summary: string
  findings: InsightFinding[]
}

export interface WorkbookInsight {
  summary: string
  sheets: SheetInsight[]
  suggestions: string[]
}

/**
 * Heuristic data insight report (ExcelGenius2-style "upload -> summary +
 * anomalies"): per-sheet one-liner, missing/duplicate/outlier/normalization
 * findings, and concrete next-step suggestions. Deterministic, no LLM needed.
 */
export async function buildWorkbookInsight(path: string, sheet?: string): Promise<WorkbookInsight> {
  const profile = await profileWorkbook(path, sheet)
  const sheetInsights = profile.sheets.map((entry) => insightForSheet(entry))
  const findings = sheetInsights.flatMap((entry) => entry.findings)
  const suggestions = buildSuggestions(profile, findings)
  return {
    summary: `共 ${profile.sheetCount} 个工作表；${findings.filter((f) => f.severity === 'alert').length} 个重点、${findings.filter((f) => f.severity === 'warn').length} 个提示。`,
    sheets: sheetInsights,
    suggestions,
  }
}

function insightForSheet(sheet: SheetProfile): SheetInsight {
  const findings: InsightFinding[] = []
  const headers = sheet.columns.filter((column) => column.header).map((column) => column.header)
  let summary = `${sheet.sheet}：${sheet.dataRows} 行数据，${sheet.columnCount} 列`
  if (headers.length > 0) summary += `，表头：${headers.slice(0, 6).join(' / ')}${headers.length > 6 ? ' …' : ''}`
  if (sheet.formulaCells > 0) summary += `，含 ${sheet.formulaCells} 个公式`

  if (sheet.dataRows === 0) {
    findings.push({ severity: 'info', category: 'empty', message: `${sheet.sheet} 没有数据行，只有表头。` })
  }
  for (const column of sheet.columns) {
    if (!column.header) continue
    const label = `${sheet.sheet}!${column.column}（${column.header}）`
    if (column.missing > 0 && column.nonEmpty + column.missing > 0) {
      const ratio = column.missing / (column.nonEmpty + column.missing)
      findings.push({
        severity: ratio >= 0.2 ? 'warn' : 'info',
        category: 'missing',
        message: `${label} 有 ${column.missing} 个空值（${Math.round(ratio * 100)}%）。`,
      })
    }
    if (column.nonEmpty > 3 && column.uniqueCapped) {
      findings.push({
        severity: 'warn',
        category: 'duplicate',
        message: `${label} 值分布很集中，疑似存在大量重复值。`,
      })
    }
    if (column.dtype === 'number' && column.mean !== undefined && column.max !== undefined && column.min !== undefined && column.mean > 0) {
      if (column.max > column.mean * 5 && column.max - column.mean > column.mean) {
        findings.push({
          severity: 'warn',
          category: 'outlier',
          message: `${label} 最大值 ${column.max} 远高于均值 ${column.mean}，疑似存在异常大值。`,
        })
      }
      if (column.min < 0 && /(金额|amount|price|cost|revenue|sales|profit|总额|费用)/i.test(column.header)) {
        findings.push({
          severity: 'warn',
          category: 'negative',
          message: `${label} 出现负数（最小值 ${column.min}），请确认是否为退款/冲销。`,
        })
      }
    }
    if (column.dtype === 'string' && column.samples.some((sample) => sample !== sample.trim())) {
      findings.push({
        severity: 'info',
        category: 'whitespace',
        message: `${label} 存在首尾空格，建议 trimText。`,
      })
    }
  }
  if (sheet.formulaCells > 0) {
    findings.push({
      severity: 'info',
      category: 'formula',
      message: `${sheet.sheet} 含 ${sheet.formulaCells} 个公式，可运行 excel_autofix 体检并修复。`,
    })
  }
  return { sheet: sheet.sheet, summary, findings }
}

function buildSuggestions(profile: WorkbookProfile, findings: InsightFinding[]): string[] {
  const suggestions: string[] = []
  if (findings.some((f) => f.category === 'missing')) suggestions.push('有缺失值：用 excel_operate 的 fillMissing 补空值，或删除整空行。')
  if (findings.some((f) => f.category === 'duplicate')) suggestions.push('疑似重复：用 dedupeRows 按关键列去重。')
  if (findings.some((f) => f.category === 'outlier' || f.category === 'negative')) suggestions.push('发现异常/负值：建议先核对源数据，再用条件格式或图表突出展示。')
  if (findings.some((f) => f.category === 'whitespace')) suggestions.push('存在首尾空格：用 trimText 清理，再去做匹配/去重。')
  if (profile.sheets.some((s) => s.formulaCells > 0)) suggestions.push('表里含公式：可运行 excel_autofix 体检并修复。')
  if (profile.sheets.some((s) => s.dataRows > 20)) suggestions.push('数据量较大：可用 excel_create_pivot / aggregateReport 做透视汇总，或 excel_create_chart 画图。')
  if (suggestions.length === 0) suggestions.push('未发现明显数据问题；可继续做报表（report）、透视或图表。')
  return suggestions
}
