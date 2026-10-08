import { readFile } from 'node:fs/promises'
import { t, listJoin, listSeparator } from './i18n.ts'
import ExcelJS from 'exceljs'
import { parseFormula, type ParsedRef } from './formula.ts'
import { stripPivotTableParts } from './workbook.ts'

const FUNCTION_DESCRIPTIONS: Record<string, string> = {
  SUM: '求和',
  AVERAGE: '求平均',
  COUNT: '计数（只数数字）',
  COUNTA: '计数（非空）',
  MAX: '取最大值',
  MIN: '取最小值',
  MEDIAN: '取中位数',
  PRODUCT: '求乘积',
  SUMIF: '按条件求和',
  SUMIFS: '按多个条件求和',
  COUNTIF: '按条件计数',
  COUNTIFS: '按多个条件计数',
  AVERAGEIF: '按条件求平均',
  AVERAGEIFS: '按多个条件求平均',
  SUBTOTAL: '分类汇总',
  VLOOKUP: '纵向查找（按首列找并返回指定列）',
  XLOOKUP: '查找并返回匹配值',
  INDEX: '按行列位置取值',
  MATCH: '查找目标所在位置',
  IF: '条件判断，成立返回一个值、否则返回另一个值',
  IFERROR: '出错时返回替代值',
  IFNA: '查不到时返回替代值',
  TODAY: '返回当天日期',
  NOW: '返回当前日期时间',
  YEAR: '取年份',
  MONTH: '取月份',
  DAY: '取日',
  DATE: '按年月日拼日期',
  DATEDIF: '计算两个日期的间隔',
  EOMONTH: '返回某月最后一天',
  TEXT: '按格式转文本',
  ROUND: '四舍五入',
  ROUNDUP: '向上取整',
  ROUNDDOWN: '向下取整',
  INT: '取整',
  MOD: '取余数',
  ABS: '取绝对值',
  TRIM: '去掉多余空格',
  LEN: '计算字符数',
  LEFT: '取左侧若干个字符',
  RIGHT: '取右侧若干个字符',
  MID: '从中间取字符',
  CONCAT: '拼接文本',
  CONCATENATE: '拼接文本',
  SUBSTITUTE: '替换文本',
}

/**
 * Translate a function description at lookup time. Calling `t()` inside the map
 * above would freeze the language at import, before the plugin config has had a
 * chance to set it — the map is module-level and evaluated on first import.
 */
function describeFunction(name: string): string {
  return t(FUNCTION_DESCRIPTIONS[name] ?? name)
}

export interface FormulaExplanation {
  formula: string
  summary: string
  details: string[]
  references: string[]
}

/** Explain an Excel formula in plain language (cellm/xeli-style). */
export function explainFormula(formula: string): FormulaExplanation {
  const parsed = parseFormula(formula)
  const raw = formula.trim().replace(/^=/, '')
  const functions = [...raw.matchAll(/([A-Za-z]+)\s*\(/g)]
    .map((match) => match[1]!.toUpperCase())
    .filter((name) => FUNCTION_DESCRIPTIONS[name])
  const references = parsed.references.map(refText)
  const details: string[] = []
  if (functions.length > 0) {
    details.push(t('使用了函数：{list}', { list: listJoin(functions.map((name) => t('{name}（{description}）', { name, description: describeFunction(name) }))) }))
  }
  if (references.length > 0) {
    details.push(t('引用区域：{list}', { list: listJoin(references) }))
    const crossSheet = parsed.references.filter((ref) => ref.start.sheet)
    if (crossSheet.length > 0) {
      details.push(t('涉及跨表引用：{list}', { list: listJoin([...new Set(crossSheet.map((ref) => ref.start.sheet))]) }))
    }
  }
  if (/[+\-*/^]/.test(raw)) details.push(t('包含算术运算（加/减/乘/除/乘方）'))
  if (/[<>=]/.test(raw.replace(/=+/g, '='))) details.push(t('包含比较判断'))
  const summary = functions.length > 0
    ? t('这是一个 {functions} 公式：{descriptions}。', { functions: functions.join(' + '), descriptions: listJoin(functions.map(describeFunction), 'semicolon') })
    : references.length > 0
      ? t('这是一个引用其他单元格/区域参与计算或比较的公式。')
      : t('这是一个常量或简单表达式。')
  return { formula: formula.trim(), summary, details, references }
}

/** Read the formula (or value) of one cell from an .xlsx file. */
export async function readCellContent(path: string, cellId: string): Promise<string> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(stripPivotTableParts(await readFile(path)) as any)
  const bang = cellId.lastIndexOf('!')
  const sheetName = bang >= 0 ? cellId.slice(0, bang) : workbook.worksheets[0]!.name
  const address = bang >= 0 ? cellId.slice(bang + 1) : cellId
  const sheet = workbook.worksheets.find((entry) => entry.name.toLowerCase() === sheetName.toLowerCase())
  if (!sheet) throw new Error(`sheet not found: ${sheetName}`)
  const cell = sheet.getCell(address)
  if (cell.formula) return `=${cell.formula}`
  const value = cell.value
  if (value === null || value === undefined) return ''
  return typeof value === 'object' ? JSON.stringify(value) : String(value)
}

function refText(ref: ParsedRef): string {
  const sheet = ref.start.sheet ? `${ref.start.sheet}!` : ''
  const start = `${ref.start.absColumn ? '$' : ''}${ref.start.column}${ref.start.absRow ? '$' : ''}${ref.start.row ?? ''}`
  if (!ref.end) return `${sheet}${start}`
  const end = `${ref.end.absColumn ? '$' : ''}${ref.end.column}${ref.end.absRow ? '$' : ''}${ref.end.row ?? ''}`
  return `${sheet}${start}:${end}`
}
