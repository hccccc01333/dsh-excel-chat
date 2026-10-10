/**
 * Inserting and deleting rows and columns, and keeping every formula honest.
 *
 * Excel shifts references when rows are inserted or deleted; nothing else in the
 * workbook does. These helpers walk the formulas and move the references, and mark
 * the ones that pointed into a deleted range so they read as errors instead of
 * silently pointing somewhere new.
 */
import ExcelJS from 'exceljs'
import { t } from '../i18n.ts'
import { columnToNumber, normalizeSheet, numberToColumn, parseCellId, parseFormula, type RefPoint } from '../formula.ts'
import type { OperationWarning } from '../operation-types.ts'
import { absoluteColumnRef, cellContentOf, findSheet, parseRange, resolveCell, shiftFormulaReferences, writeContent } from './core.ts'

/** Delete rows with the same reference-shift semantics as the deleteRows op. */
export function deleteRowsFromSheet(
  workbook: ExcelJS.Workbook,
  sheetName: string,
  start: number,
  count: number,
  warnings: OperationWarning[],
  opIndex: number,
): void {
  const sheet = findSheet(workbook, sheetName)
  if (!sheet) throw new Error(`sheet not found: ${sheetName}`)
  if (start < 1 || count < 1) throw new Error(`invalid deleteRows: row=${start} count=${count}`)
  const end = start + count - 1
  for (const formulaCell of collectDeletedRangeRefs(workbook, sheetName, start, end)) {
    warnings.push({ op: opIndex, message: t('公式 {cell} 引用了 {sheet} 中被删除的行', { cell: formulaCell, sheet: sheetName }) })
  }
  markDeletedRowRefs(workbook, sheetName, start, end)
  sheet.spliceRows(start, count)
  shiftWorkbookRows(workbook, sheetName, end + 1, -count)
}

/** Delete columns with the same reference-shift semantics as the deleteColumns op. */
export function deleteColumnsFromSheet(
  workbook: ExcelJS.Workbook,
  sheetName: string,
  column: number,
  count: number,
  warnings: OperationWarning[],
  opIndex: number,
): void {
  const sheet = findSheet(workbook, sheetName)
  if (!sheet) throw new Error(`sheet not found: ${sheetName}`)
  if (column < 1 || count < 1) throw new Error(`invalid deleteColumns: column=${column} count=${count}`)
  const end = column + count - 1
  for (const formulaCell of collectDeletedColumnRefs(workbook, sheetName, column, end)) {
    warnings.push({ op: opIndex, message: t('公式 {cell} 引用了 {sheet} 中被删除的列', { cell: formulaCell, sheet: sheetName }) })
  }
  markDeletedColumnRefs(workbook, sheetName, column, end)
  sheet.spliceColumns(column, count)
  shiftWorkbookColumns(workbook, sheetName, end + 1, -count)
}

export function shiftWorkbookRows(workbook: ExcelJS.Workbook, editedSheet: string, threshold: number, rowDelta: number): void {
  const edited = normalizeSheet(editedSheet)
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!cell.formula) return
        const formula = `=${cell.formula}`
        const shifted = shiftFormulaReferences(formula, sheet.name, edited, { rowDelta, rowThreshold: threshold })
        if (shifted !== formula) cell.value = { formula: shifted.slice(1) }
      })
    })
  })
}

export function collectDeletedRangeRefs(
  workbook: ExcelJS.Workbook,
  editedSheet: string,
  start: number,
  end: number,
): string[] {
  const edited = normalizeSheet(editedSheet)
  const hits: string[] = []
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!cell.formula) return
        const parsed = parseFormula(`=${cell.formula}`)
        for (const ref of parsed.references) {
          for (const point of [ref.start, ref.end].filter((p): p is RefPoint => p !== null)) {
            const target = normalizeSheet(point.sheet ?? sheet.name)
            if (target === edited && point.row !== null && !point.absRow && point.row >= start && point.row <= end) {
              hits.push(`${sheet.name}!${cell.address}`)
              return
            }
          }
        }
      })
    })
  })
  return hits
}

export function shiftWorkbookColumns(workbook: ExcelJS.Workbook, editedSheet: string, threshold: number, colDelta: number): void {
  const edited = normalizeSheet(editedSheet)
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!cell.formula) return
        const formula = `=${cell.formula}`
        const shifted = shiftFormulaReferences(formula, sheet.name, edited, { colDelta, colThreshold: threshold })
        if (shifted !== formula) cell.value = { formula: shifted.slice(1) }
      })
    })
  })
}

export function collectDeletedColumnRefs(
  workbook: ExcelJS.Workbook,
  editedSheet: string,
  start: number,
  end: number,
): string[] {
  const edited = normalizeSheet(editedSheet)
  const hits: string[] = []
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!cell.formula) return
        const parsed = parseFormula(`=${cell.formula}`)
        for (const ref of parsed.references) {
          for (const point of [ref.start, ref.end].filter((p): p is RefPoint => p !== null)) {
            const target = normalizeSheet(point.sheet ?? sheet.name)
            const columnNumber = columnToNumber(point.column)
            if (target === edited && !point.absColumn && columnNumber >= start && columnNumber <= end) {
              hits.push(`${sheet.name}!${cell.address}`)
              return
            }
          }
        }
      })
    })
  })
  return hits
}

export function markDeletedRowRefs(workbook: ExcelJS.Workbook, editedSheet: string, start: number, end: number): void {
  const edited = normalizeSheet(editedSheet)
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!cell.formula) return
        const formula = `=${cell.formula}`
        const rewritten = shiftFormulaReferences(formula, sheet.name, edited, {
          rowDeletedStart: start,
          rowDeletedEnd: end,
        })
        if (rewritten !== formula) cell.value = { formula: rewritten.slice(1) }
      })
    })
  })
}

export function markDeletedColumnRefs(workbook: ExcelJS.Workbook, editedSheet: string, start: number, end: number): void {
  const edited = normalizeSheet(editedSheet)
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!cell.formula) return
        const formula = `=${cell.formula}`
        const rewritten = shiftFormulaReferences(formula, sheet.name, edited, {
          colDeletedStart: start,
          colDeletedEnd: end,
        })
        if (rewritten !== formula) cell.value = { formula: rewritten.slice(1) }
      })
    })
  })
}
