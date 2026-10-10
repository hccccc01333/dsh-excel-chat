/**
 * Joining two sheets on a key column.
 *
 * The join is exact-match and writes values, not formulas — a VLOOKUP would break
 * the moment the lookup sheet is sorted, and the point of this operation is a
 * filled-in column that stays put. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs'
import { t } from '../i18n.ts'
import { columnToNumber, numberToColumn } from '../formula.ts'
import { cellContent, contentToCellValue } from '../workbook.ts'
import type { ExcelOperation, OperationWarning } from '../operation-types.ts'
import { cellContentOf, findSheet, parseRange, resolveCell, writeContent } from './core.ts'

export function joinSheets(
  workbook: ExcelJS.Workbook,
  operation: Extract<ExcelOperation, { op: 'joinSheets' }>,
  warnings: OperationWarning[],
  opIndex: number,
): void {
  if (operation.valueColumns.length !== operation.outputColumns.length) {
    throw new Error(`joinSheets valueColumns (${operation.valueColumns.length}) and outputColumns (${operation.outputColumns.length}) must have the same length`)
  }
  const sourceParsed = parseRange(workbook, operation.source)
  const lookupParsed = parseRange(workbook, operation.lookup)
  const lookupKeyCol = columnToNumber(operation.lookupKey)

  // First match wins, mirroring VLOOKUP's approximate=false behaviour.
  const index = new Map<string, string[]>()
  for (let row = lookupParsed.startRow + 1; row <= lookupParsed.endRow; row++) {
    const key = normalizeJoinKey(lookupParsed.sheet.getCell(`${numberToColumn(lookupKeyCol)}${row}`).value)
    if (!key || index.has(key)) continue
    index.set(key, operation.valueColumns.map((column) =>
      cellContentOf(lookupParsed.sheet.getCell(`${numberToColumn(columnToNumber(column))}${row}`)),
    ))
  }

  const sourceKeyCol = columnToNumber(operation.sourceKey)
  let matched = 0
  let missed = 0
  for (let row = sourceParsed.startRow + 1; row <= sourceParsed.endRow; row++) {
    const key = normalizeJoinKey(sourceParsed.sheet.getCell(`${numberToColumn(sourceKeyCol)}${row}`).value)
    const values = key ? index.get(key) : undefined
    if (!values) {
      missed++
      if (operation.missValue !== undefined) {
        operation.outputColumns.forEach((column, i) => {
          sourceParsed.sheet.getCell(`${numberToColumn(columnToNumber(column))}${row}`).value =
            typeof operation.missValue === 'number' ? operation.missValue : String(operation.missValue ?? '')
        })
      }
      continue
    }
    matched++
    values.forEach((value, i) => {
      const column = columnToNumber(operation.outputColumns[i]!)
      sourceParsed.sheet.getCell(`${numberToColumn(column)}${row}`).value =
        value.startsWith('=') ? { formula: value.slice(1) } : contentToCellValue(value)
    })
  }
  warnings.push({ op: opIndex, message: t('joinSheets 匹配了 {matched} 行，{missed} 行没有找到对应值', { matched, missed }) })
}


function normalizeJoinKey(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return ''
  return String(typeof value === 'object' && !(value instanceof Date) ? JSON.stringify(value) : value).trim().toLowerCase()
}
