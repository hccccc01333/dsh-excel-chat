import ExcelJS from 'exceljs'
import { stripPivotTableParts } from './workbook.ts'

export interface ErrorCell {
  /** Sheet-qualified cell id, e.g. "订单!D4". */
  cell: string
  /** The error value, e.g. "#DIV/0!". */
  error: string
  /** The formula that produced it, when the cell holds one. */
  formula: string | null
}

export interface ErrorScan {
  total: number
  /** Per-error-code counts, e.g. { "#DIV/0!": 2, "#N/A": 1 }. */
  counts: Record<string, number>
  errorCells: ErrorCell[]
  sheetsScanned: string[]
}

/** The error values Excel can store in a cell. */
export const EXCEL_ERROR_VALUES = [
  '#DIV/0!',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#REF!',
  '#VALUE!',
  '#GETTING_DATA',
] as const

/**
 * Find cells whose value is an Excel *error* rather than text that merely looks
 * like one.
 *
 * This reads the workbook directly instead of going through readWorkbookCells:
 * that helper stringifies every value, which would make a literal "#N/A" typed
 * by a user indistinguishable from a genuine error — and flagging those would
 * turn the check into noise.
 */
export async function findErrorCells(data: Uint8Array, sheetFilter?: string): Promise<ErrorScan> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(stripPivotTableParts(data) as any)
  const errorCells: ErrorCell[] = []
  const counts: Record<string, number> = {}
  const sheetsScanned: string[] = []
  workbook.eachSheet((sheet) => {
    if (sheetFilter !== undefined && sheet.name !== sheetFilter) return
    sheetsScanned.push(sheet.name)
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const error = errorValueOf(cell)
        if (error === null) return
        counts[error] = (counts[error] ?? 0) + 1
        errorCells.push({
          cell: `${sheet.name}!${cell.address}`,
          error,
          formula: cell.formula ? `=${cell.formula}` : null,
        })
      })
    })
  })
  return { total: errorCells.length, counts, errorCells, sheetsScanned }
}

/**
 * The error value a cell holds, if any.
 *
 * A bare error cell reports `type === Error` with `{ error }`, but a *formula*
 * whose cached result is an error keeps `type === Formula` and buries the error
 * in `value.result` — checking only for the Error type would miss every
 * formula-produced error, which is exactly the interesting case. Reading the
 * value shape covers both without depending on the type enum.
 */
function errorValueOf(cell: ExcelJS.Cell): string | null {
  const value = cell.value as { error?: unknown; result?: { error?: unknown } } | null
  if (value === null || typeof value !== 'object') return null
  if (typeof value.error === 'string') return value.error
  if (value.result !== null && typeof value.result === 'object' && typeof value.result.error === 'string') {
    return value.result.error
  }
  return null
}
