import ExcelJS from 'exceljs'
import { readFile } from 'node:fs/promises'
import { columnToNumber, numberToColumn } from './formula.ts'

export interface ReadCell {
  id: string
  value: string | number | boolean | null
  formula?: string
  type: 'string' | 'number' | 'boolean' | 'date' | 'formula' | 'empty'
  numberFormat?: string
  bold?: boolean
  italic?: boolean
  fontSize?: number
  fontName?: string
  fontColor?: string
  fill?: string
  hAlign?: string
  vAlign?: string
  wrapText?: boolean
  mergedTo?: string
  dataValidationType?: string
}

export interface ReadSheetResult {
  sheet: string
  range: string
  cells: ReadCell[]
}

export interface ReadWorkbookOptions {
  sheet?: string
  /** A1 range on the selected sheet, e.g. "A1:D20". */
  range?: string
  cells?: string[]
}

function colorToHex(color: Partial<ExcelJS.Color> | undefined): string | undefined {
  if (!color || typeof color.argb !== 'string') return undefined
  const argb = color.argb.toUpperCase()
  return argb.startsWith('FF') ? argb.slice(2) : argb
}

function describeValue(cell: ExcelJS.Cell): ReadCell['value'] {
  if (cell.formula) return null
  const value = cell.value
  if (value === null || value === undefined) return null
  if (typeof value === 'object') {
    if (value instanceof Date) return value.toISOString()
    if ('error' in value) return String(value.error)
    const text = (value as { text?: unknown }).text
    if (typeof text === 'string') return text
    return JSON.stringify(value)
  }
  return value
}

function describeCell(sheet: ExcelJS.Worksheet, column: string, row: number): ReadCell {
  const cell = sheet.getCell(`${column}${row}`)
  const raw = cell.value
  const isDate = raw instanceof Date
  const isFormula = Boolean(cell.formula)
  const isEmpty = raw === null || raw === undefined
  const type: ReadCell['type'] = isFormula ? 'formula' : isEmpty ? 'empty' : isDate ? 'date' : typeof raw === 'number' ? 'number' : typeof raw === 'boolean' ? 'boolean' : 'string'
  const font = cell.font ?? {}
  const fill = cell.fill
  const merges = (sheet.model as { merges?: string[] }).merges ?? []
  const mergedTo = cell.isMerged ? merges.find((merge) => merge.includes(`${column}${row}`)) : undefined
  return {
    id: `${sheet.name}!${column}${row}`,
    value: describeValue(cell),
    formula: isFormula ? `=${cell.formula}` : undefined,
    type,
    numberFormat: cell.numFmt && cell.numFmt !== 'General' ? cell.numFmt : undefined,
    bold: font.bold ?? undefined,
    italic: font.italic ?? undefined,
    fontSize: font.size ?? undefined,
    fontName: font.name ?? undefined,
    fontColor: colorToHex(font.color),
    fill: fill && fill.type === 'pattern' && (fill as { fgColor?: Partial<ExcelJS.Color> }).fgColor
      ? colorToHex((fill as { fgColor?: Partial<ExcelJS.Color> }).fgColor)
      : undefined,
    hAlign: cell.alignment?.horizontal ?? undefined,
    vAlign: cell.alignment?.vertical ?? undefined,
    wrapText: cell.alignment?.wrapText ?? undefined,
    mergedTo,
    dataValidationType: cell.dataValidation?.type,
  }
}

/** Precisely read cells (values, formulas, types, and formats) from an .xlsx file. */
export async function readWorkbookDetail(path: string, options: ReadWorkbookOptions = {}): Promise<ReadSheetResult[]> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await readFile(path) as any)
  const results: ReadSheetResult[] = []
  for (const sheet of workbook.worksheets) {
    if (options.sheet && sheet.name.toLowerCase() !== options.sheet.toLowerCase()) continue
    const parsed = options.range ? parseRangeText(options.range) : null
    const startCol = parsed?.startCol ?? 1
    const startRow = parsed?.startRow ?? 1
    const endCol = parsed?.endCol ?? sheet.columnCount
    const endRow = parsed?.endRow ?? sheet.rowCount
    const requested = options.cells ? new Set(options.cells.map((id) => `${id.split('!').pop()?.toUpperCase()}`)) : null
    const cells: ReadCell[] = []
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const column = numberToColumn(col)
        if (requested && !requested.has(`${column}${row}`)) continue
        const cell = describeCell(sheet, column, row)
        if (cell.type === 'empty' && !requested) continue
        cells.push(cell)
      }
    }
    results.push({
      sheet: sheet.name,
      range: `${numberToColumn(startCol)}${startRow}:${numberToColumn(endCol)}${endRow}`,
      cells,
    })
  }
  return results
}

function parseRangeText(range: string): { startCol: number; startRow: number; endCol: number; endRow: number } {
  const body = range.includes('!') ? range.slice(range.lastIndexOf('!') + 1) : range
  const match = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(body)
  if (!match) throw new Error(`invalid range: ${range}`)
  const endCol = match[3] ? columnToNumber(match[3]!) : columnToNumber(match[1]!)
  const endRow = match[4] ? Number(match[4]!) : Number(match[2]!)
  return { startCol: columnToNumber(match[1]!), startRow: Number(match[2]!), endCol, endRow }
}
