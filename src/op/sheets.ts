/**
 * Whole-sheet concerns: visibility, order, tab colour, hyperlinks, freeze panes.
 *
 * Everything here acts on the sheet as an object rather than on its cells, which is
 * why it reads as one group. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs'
import { columnToNumber, numberToColumn } from '../formula.ts'
import { t } from '../i18n.ts'
import { cellContent } from '../workbook.ts'
import type { ExcelOperation } from '../operation-types.ts'
import { absoluteColumnRef, findSheet, parseRange, parseTargetCell, qualifySheetName, resolveCell } from './core.ts'

export function setHyperlink(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, { op: 'setHyperlink' }>): void {
  const cell = resolveCell(workbook, options.cell)
  const text = options.text
  if (options.url) {
    cell.value = { text: text ?? options.url, hyperlink: options.url }
    return
  }
  if (options.location) {
    // ExcelJS mishandles internal links (writes both a "#..." location AND an
    // External relationship), so use a HYPERLINK() formula, which Excel always
    // navigates correctly. The location's sheet name is quoted if needed.
    const loc = internalLinkLocation(options.location)
    const label = text ?? options.location.replace(/^#/, '')
    cell.value = { formula: `HYPERLINK("${loc}","${label.replaceAll('"', '""')}")` }
    return
  }
  throw new Error('setHyperlink requires url (external) or location (internal, e.g. "Sheet2!A1")')
}


function internalLinkLocation(location: string): string {
  const bare = location.startsWith('#') ? location.slice(1) : location
  const bang = bare.lastIndexOf('!')
  if (bang < 0) return `#${bare}` // a defined name
  const sheet = bare.slice(0, bang).replaceAll(/^'|'$/g, '')
  const ref = bare.slice(bang + 1)
  return `#${qualifySheetName(sheet)}!${ref}`
}


export function copyStyle(workbook: ExcelJS.Workbook, sourceId: string, targetRange: string): void {
  const source = resolveCell(workbook, sourceId)
  const style: ExcelJS.Style = JSON.parse(JSON.stringify(source.style ?? {}))
  const parsed = parseRange(workbook, targetRange)
  for (let row = parsed.startRow; row <= parsed.endRow; row++) {
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`)
      cell.style = JSON.parse(JSON.stringify(style))
    }
  }
}


export function freezeFormulas(workbook: ExcelJS.Workbook, range: string): { frozen: number; skipped: number } {
  const parsed = parseRange(workbook, range)
  let frozen = 0
  let skipped = 0
  for (let row = parsed.startRow; row <= parsed.endRow; row++) {
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`)
      if (!cell.formula) continue
      // Plugin-written formulas often carry no cached result (never opened in
      // Excel). Freezing those to "null" would erase them, so leave them as
      // formulas and skip.
      const result = cell.result
      if (result === undefined || result === null) {
        skipped++
        continue
      }
      cell.value = result
      frozen++
    }
  }
  return { frozen, skipped }
}


export function uniqueValues(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, { op: 'uniqueValues' }>): number {
  const parsed = parseRange(workbook, options.source)
  const target = parseTargetCell(workbook, options.target, parsed.sheet.name)
  const seen = new Set<string>()
  const ordered: ExcelJS.CellValue[] = []
  const firstDataRow = options.includeHeader ? parsed.startRow : parsed.startRow + 1
  for (let row = firstDataRow; row <= parsed.endRow; row++) {
    const cell = parsed.sheet.getCell(`${numberToColumn(parsed.startCol)}${row}`)
    const raw = cell.formula ? cell.result : cell.value
    const key = uniqueValueKey(cell, raw)
    if (seen.has(key)) continue
    seen.add(key)
    ordered.push(raw)
  }
  let outRow = target.row
  for (const value of ordered) {
    target.sheet.getCell(`${numberToColumn(target.col)}${outRow}`).value = value === null || value === undefined ? '' : value
    outRow++
  }
  return ordered.length
}


function uniqueValueKey(cell: ExcelJS.Cell, raw: ExcelJS.CellValue): string {
  if (cell.formula && (raw === undefined || raw === null)) return `=f:${cell.formula}`
  if (raw === null || raw === undefined) return '∅'
  if (raw instanceof Date) return `date:${raw.getTime()}`
  if (typeof raw === 'object') return `obj:${JSON.stringify(raw)}`
  return `${typeof raw}:${String(raw)}`
}


export function applySheetView(sheet: ExcelJS.Worksheet, patch: (view: ExcelJS.WorksheetView) => void): void {
  // exceljs types views as strict unions but accepts partial views at runtime,
  // so the default view is cast from a minimal object.
  const existing = (sheet.views ?? []) as ExcelJS.WorksheetView[]
  const views: ExcelJS.WorksheetView[] = existing.length
    ? existing
    : [{ workbookViewId: 0 } as unknown as ExcelJS.WorksheetView]
  for (const view of views) patch(view)
  sheet.views = views as ExcelJS.WorksheetView[]
}


export function moveSheet(workbook: ExcelJS.Workbook, name: string, position: number): void {
  const sheet = findSheet(workbook, name)
  if (!sheet) throw new Error(`sheet not found: ${name}`)
  const ordered = workbook.worksheets
  const others = ordered.filter((entry) => entry.id !== sheet.id)
  const clamped = Math.max(1, Math.min(position, ordered.length))
  const before = others.slice(0, clamped - 1)
  const after = others.slice(clamped - 1)
  ;[...before, sheet, ...after].forEach((entry, i) => {
    ;(entry as ExcelJS.Worksheet & { orderNo: number }).orderNo = i + 1
  })
}


export function applyRankColumn(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, { op: 'rankColumn' }>): void {
  const parsed = parseRange(workbook, options.range)
  const metricCol = columnToNumber(options.metricColumn)
  if (metricCol < parsed.startCol || metricCol > parsed.endCol) {
    throw new Error(`rankColumn metric column outside range: ${options.metricColumn}`)
  }
  const firstData = options.skipHeader === false ? parsed.startRow : parsed.startRow + 1
  const metricRange = absoluteColumnRef(parsed.sheet.name, numberToColumn(metricCol), firstData, parsed.endRow)
  for (let row = firstData; row <= parsed.endRow; row++) {
    parsed.sheet.getCell(`${options.outputColumn}${row}`).value = {
      formula: `RANK(${numberToColumn(metricCol)}${row},${metricRange},${options.descending === false ? 1 : 0})`,
    }
  }
}
