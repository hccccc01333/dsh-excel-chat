/**
 * Page setup and the text metrics that drive column widths.
 *
 * `displayWidth` counts CJK characters as two columns, which is what makes an
 * autofit look right in a sheet full of Chinese; it lives with page setup because
 * both are about how the sheet presents rather than what it contains. Split out of
 * `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs'
import { columnToNumber } from '../formula.ts'
import { findSheet } from './core.ts'
import type { ExcelOperation } from '../operation-types.ts'

export function applyPageSetup(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'pageSetup' }>,
): void {
  const sheet = findSheet(workbook, options.sheet)
  if (!sheet) throw new Error(`sheet not found: ${options.sheet}`)
  const pageSetup = sheet.pageSetup
  if (options.printArea) pageSetup.printArea = options.printArea
  if (options.orientation) pageSetup.orientation = options.orientation
  if (options.fitToPage !== undefined) pageSetup.fitToPage = options.fitToPage
  if (options.fitToWidth !== undefined) pageSetup.fitToWidth = options.fitToWidth
  if (options.fitToHeight !== undefined) pageSetup.fitToHeight = options.fitToHeight
  if (options.margins) pageSetup.margins = { ...pageSetup.margins, ...options.margins }
  if (options.centerHorizontally !== undefined) pageSetup.horizontalCentered = options.centerHorizontally
  if (options.centerVertically !== undefined) pageSetup.verticalCentered = options.centerVertically
}









export function displayTextOf(cell: ExcelJS.Cell): string {
  const value = cell.formula ? cell.result : cell.value
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return '2026-12-31'
  if (typeof value === 'object') return JSON.stringify(value)
  return String(value)
}


export function displayWidth(text: string): number {
  let width = 0
  for (const char of text) {
    const code = char.codePointAt(0) ?? 0
    width += code > 0x2e7f ? 2 : 1
  }
  return width
}
