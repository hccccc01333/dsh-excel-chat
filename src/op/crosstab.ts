/**
 * The two-dimensional cross-tab: row dimension × column dimension, cells live SUMIFS.
 *
 * Unlike a pivot table, every cell here is a formula, so the result updates when the
 * source does. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs'
import { columnToNumber, numberToColumn } from '../formula.ts'
import { t } from '../i18n.ts'
import { cellContent } from '../workbook.ts'
import { writeWorkbookSafely } from '../safe-write.ts'
import type { ExcelOperation, OperationWarning } from '../operation-types.ts'
import { absoluteColumnRef, findSheet, parseRange } from './core.ts'

const CROSSTAB_FUNCTIONS: Record<string, string> = {
  sum: 'SUMIFS',
  average: 'AVERAGEIFS',
  count: 'COUNTIFS',
  counta: 'COUNTIFS',
  max: 'MAXIFS',
  min: 'MINIFS',
}


const CROSSTAB_TOTALABLE = new Set(['sum', 'count', 'counta'])

export function applyCrosstab(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'crosstab' }>,
  warnings: OperationWarning[],
  opIndex: number,
): void {
  const needsMetric = options.metric.function !== 'count' && options.metric.function !== 'counta'
  if (needsMetric && !options.metric.column) {
    throw new Error(`crosstab function "${options.metric.function}" requires metric.column`)
  }
  const parsed = parseRange(workbook, options.source)
  const rowCol = columnToNumber(options.rowColumn)
  const colCol = columnToNumber(options.columnColumn)
  const firstData = parsed.startRow + 1

  // Keep the raw cell value so the output header/label cells match the source
  // criteria (dates as serials, numbers as numbers); text is only for dedup.
  const collectKeys = (col: number): Array<{ raw: ExcelJS.CellValue; text: string }> => {
    const keys: Array<{ raw: ExcelJS.CellValue; text: string }> = []
    const seen = new Set<string>()
    for (let row = firstData; row <= parsed.endRow; row++) {
      const raw = parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value
      const text = raw === null || raw === undefined ? '' : String(raw)
      if (!seen.has(text)) {
        seen.add(text)
        keys.push({ raw, text })
      }
    }
    return keys
  }
  const rowKeys = collectKeys(rowCol)
  const colKeys = collectKeys(colCol)
  if (!rowKeys.length || !colKeys.length) throw new Error('crosstab source has no data rows')

  const sheetRange = (col: number): string =>
    absoluteColumnRef(parsed.sheet.name, numberToColumn(col), firstData, parsed.endRow)
  const rowRange = sheetRange(rowCol)
  const colRange = sheetRange(colCol)
  // SUMIFS/AVERAGEIFS/MAXIFS/MINIFS take a leading sum/average range; the
  // *IFS count form takes only criteria pairs, so metric.range must be dropped
  // there or the argument list becomes invalid (odd count of range/criteria).
  const usesMetric = options.metric.function !== 'count' && options.metric.function !== 'counta'
  const metricRange = usesMetric ? sheetRange(columnToNumber(options.metric.column!)) : null
  const fn = CROSSTAB_FUNCTIONS[options.metric.function]!

  const outputSheetName = options.outputSheet ?? `${parsed.sheet.name}-交叉表`
  let output = findSheet(workbook, outputSheetName)
  if (!output) output = workbook.addWorksheet(outputSheetName)

  const rowHeader = String(parsed.sheet.getCell(`${numberToColumn(rowCol)}${parsed.startRow}`).value ?? options.rowColumn)
  const colHeader = String(parsed.sheet.getCell(`${numberToColumn(colCol)}${parsed.startRow}`).value ?? options.columnColumn)
  const corner = output.getCell('A1')
  corner.value = `${rowHeader}\\${colHeader}`
  corner.font = { bold: true }
  colKeys.forEach((key, i) => {
    const cell = output!.getCell(`${numberToColumn(2 + i)}1`)
    cell.value = key.raw === undefined || key.raw === null ? '' : key.raw
    cell.font = { bold: true }
  })

  rowKeys.forEach((rowKey, rowIndex) => {
    const outRow = 2 + rowIndex
    output!.getCell(`A${outRow}`).value = rowKey.raw === undefined || rowKey.raw === null ? '' : rowKey.raw
    colKeys.forEach((_colKey, colIndex) => {
      const columnLetter = numberToColumn(2 + colIndex)
      // Criteria point at output-sheet cells, so keys never need quoting.
      const body = metricRange
        ? `${metricRange},${rowRange},$A${outRow},${colRange},${columnLetter}$1`
        : `${rowRange},$A${outRow},${colRange},${columnLetter}$1`
      const formula = `${fn}(${body})`
      output!.getCell(`${columnLetter}${outRow}`).value = {
        formula: options.metric.function === 'average' ? `IFERROR(${formula},0)` : formula,
      }
    })
  })

  const totals = options.totals ?? true
  if (totals && CROSSTAB_TOTALABLE.has(options.metric.function)) {
    const totalRow = 2 + rowKeys.length
    const totalCol = 2 + colKeys.length
    output!.getCell(`A${totalRow}`).value = '总计'
    output!.getCell(`A${totalRow}`).font = { bold: true }
    colKeys.forEach((_colKey, colIndex) => {
      const columnLetter = numberToColumn(2 + colIndex)
      const cell = output!.getCell(`${columnLetter}${totalRow}`)
      cell.value = { formula: `SUM(${columnLetter}2:${columnLetter}${totalRow - 1})` }
      cell.font = { bold: true }
    })
    for (let row = 2; row <= totalRow; row++) {
      const last = numberToColumn(totalCol - 1)
      const cell = output!.getCell(`${numberToColumn(totalCol)}${row}`)
      cell.value = { formula: `SUM(B${row}:${last}${row})` }
      if (row === totalRow) cell.font = { bold: true }
    }
  }
  warnings.push({
    op: opIndex,
    message: t('crosstab 在 {sheet} 生成了 {rows}x{cols} 的交叉表（{fn} 活公式）', { sheet: outputSheetName, rows: rowKeys.length, cols: colKeys.length, fn }),
  })
}

/** Formats exceljs can embed, and the size ceiling that keeps memory sane. */
/**
 * Embed an image anchored at a cell. Cross-platform: exceljs writes the media
 * part and the drawing XML itself, so no Excel installation is involved. Because
 * we mutate the loaded workbook (rather than rebuilding it), images that were
 * already in the file survive — verified by reading the media parts back.
 */
/** Resolve image bytes plus exceljs's extension token, from a path or base64. */

/**
 * The rendered size for an embedded image.
 *
 * Both sides given → use them as-is. Neither → the image's own pixel size. Just
 * one → scale the other from the intrinsic size, so asking for `width: 300` on a
 * 100×200 image renders 300×600 rather than squashing it to 300×100. Returns
 * null only when the caller gave a single side and the header is unreadable.
 */
/** Proportional counterpart to a given dimension, never rounding down to zero. */
