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
import { cellContentOf, findSheet, parseRange, resolveCell, shiftFormulaReferences, writeContent } from './core.ts'

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
  type LookupValue = { raw: ExcelJS.CellValue; formula: string | null; row: number; col: number }
  const index = new Map<string, LookupValue[]>()
  for (let row = lookupParsed.startRow + 1; row <= lookupParsed.endRow; row++) {
    const key = normalizeJoinKey(lookupParsed.sheet.getCell(`${numberToColumn(lookupKeyCol)}${row}`).value)
    if (!key || index.has(key)) continue
    index.set(key, operation.valueColumns.map((column) => {
      const col = columnToNumber(column)
      const cell = lookupParsed.sheet.getCell(`${numberToColumn(col)}${row}`)
      return { raw: cell.value, formula: cell.formula ? `=${cell.formula}` : null, row, col }
    }))
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
      const cell = sourceParsed.sheet.getCell(`${numberToColumn(column)}${row}`)
      if (value.formula) {
        // A copied formula must follow its new row, and a cross-sheet lookup
        // must gain the destination sheet as its base for unqualified refs.
        const shifted = shiftFormulaReferences(value.formula, lookupParsed.sheet.name, null, {
          rowDelta: row - value.row,
          colDelta: column - value.col,
        })
        cell.value = contentToCellValue(shifted)
      } else {
        cell.value = value.raw
      }
    })
  }
  warnings.push({ op: opIndex, message: t('joinSheets 匹配了 {matched} 行，{missed} 行没有找到对应值', { matched, missed }) })
}


function normalizeJoinKey(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return ''
  return String(typeof value === 'object' && !(value instanceof Date) ? JSON.stringify(value) : value).trim().toLowerCase()
}
