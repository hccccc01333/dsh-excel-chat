import ExcelJS from 'exceljs'
import {
  columnToNumber,
  normalizeSheet,
  numberToColumn,
  parseCellId,
  parseFormula,
  type RefPoint,
} from './formula.ts'
import { validate, type ValidationResult } from './validator.ts'
import { readWorkbookCells } from './workbook.ts'
import { readFile } from 'node:fs/promises'

export type ExcelOperation =
  | { op: 'set'; cells: Record<string, string> }
  | { op: 'fill'; source: string; target: string }
  | { op: 'insertRows'; sheet: string; row: number; count: number }
  | { op: 'deleteRows'; sheet: string; row: number; count: number }
  | { op: 'insertColumns'; sheet: string; column: string; count: number }
  | { op: 'deleteColumns'; sheet: string; column: string; count: number }
  | { op: 'addSheet'; name: string }
  | { op: 'renameSheet'; oldName: string; newName: string }
  | { op: 'deleteSheet'; name: string }
  | { op: 'clear'; cells: string[] }
  | { op: 'merge'; range: string }
  | { op: 'unmerge'; range: string }
  | { op: 'copyRange'; source: string; target: string; move?: boolean }
  | { op: 'fillSeries'; start: string; target: string; step?: number }
  | { op: 'style'; range: string; style: ExcelStyle }
  | { op: 'setColumnWidth'; sheet: string; column: string; width: number }
  | { op: 'setRowHeight'; sheet: string; row: number; height: number }
  | { op: 'freezePanes'; sheet: string; row: number; column: string }
  | { op: 'findReplace'; find: string; replace: string; sheet?: string; matchCase?: boolean }
  | { op: 'duplicateSheet'; name: string; newName: string }
  | { op: 'hideSheet'; name: string; hidden?: boolean }
  | { op: 'setTabColor'; name: string; color: string }
  | { op: 'sortRange'; range: string; keys: Array<{ column: string; direction?: 'asc' | 'desc' }>; headerRows?: number }
  | {
      op: 'dataValidation'
      range: string
      type: 'list' | 'whole' | 'decimal' | 'date' | 'textLength' | 'custom'
      operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual'
      formula1?: string
      formula2?: string
      allowBlank?: boolean
      showInputMessage?: boolean
      prompt?: string
      showErrorMessage?: boolean
      errorStyle?: 'stop' | 'warning' | 'information'
      error?: string
      errorTitle?: string
    }
  | {
      op: 'conditionalFormatting'
      range: string
      rules: Array<{
        type: 'cellIs' | 'expression'
        operator?: string
        formula?: string | number
        formula2?: string | number
        style?: ExcelStyle
      }>
    }
  | { op: 'autoFilter'; range: string }
  | { op: 'subtotal'; sheet: string; range: string; groupColumn: string; summaryColumns: Array<{ column: string; function: 'sum' | 'average' | 'count' | 'max' | 'min' }>; addGrandTotal?: boolean }
  | { op: 'aggregateReport'; source: string; groupColumn: string; metrics: Array<{ column: string; function: 'sum' | 'average' | 'count' | 'counta' | 'max' | 'min' }>; outputSheet?: string }
  | { op: 'filterToRange'; source: string; criteria: Array<{ column: string; operator: 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains'; value: string | number }>; target: string; matchAll?: boolean }
  | { op: 'protectSheet'; sheet: string; password?: string }
  | { op: 'unprotectSheet'; sheet: string; password?: string }
  | { op: 'mailMerge'; template: string; data: string; outputSheet?: string }
  | {
      op: 'addTable'
      name: string
      range: string
      headerRow?: boolean
      totalsRow?: boolean
      showRowStripes?: boolean
      showColumnStripes?: boolean
    }

export interface ExcelStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontColor?: string
  fill?: string
  numberFormat?: string
  hAlign?: 'left' | 'center' | 'right'
  vAlign?: 'top' | 'middle' | 'bottom'
  wrapText?: boolean
}

export interface OperationWarning {
  op: number
  message: string
}

export interface ApplyOperationsResult {
  warnings: OperationWarning[]
}

export interface OperateResult extends ApplyOperationsResult {
  outputPath: string
  validation: ValidationResult
}

const RANGE_LINE = /^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/

export function findSheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined {
  const normalized = normalizeSheet(name)
  return workbook.worksheets.find((sheet) => normalizeSheet(sheet.name) === normalized)
}

function resolveCell(workbook: ExcelJS.Workbook, id: string): ExcelJS.Cell {
  const parsed = parseCellId(id)
  const sheet = findSheet(workbook, parsed.sheet)
  if (!sheet) throw new Error(`sheet not found: ${parsed.sheet}`)
  return sheet.getCell(`${parsed.column}${parsed.row}`)
}

function writeContent(cell: ExcelJS.Cell, content: string): void {
  const trimmed = content.trim()
  cell.value = toCellValue(trimmed)
}

/**
 * Convert user-provided text into an Excel value: formulas stay formulas,
 * plain numbers/dates/booleans keep their type, everything else is text.
 * Workplace spreadsheets break when "100" is written as text, so numeric
 * strings are typed before they reach ExcelJS.
 */
function toCellValue(content: string): ExcelJS.CellValue {
  if (content.startsWith('=')) return { formula: content.slice(1) }
  if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(content)) return Number(content)
  if (/^true$/i.test(content)) return true
  if (/^false$/i.test(content)) return false
  const date = /^(\d{4})-(\d{2})-(\d{2})([T ](\d{2}):(\d{2})(:(\d{2}))?)?$/.exec(content)
  if (date) {
    const year = Number(date[1])
    const month = Number(date[2])
    const day = Number(date[3])
    const hour = date[5] ? Number(date[5]) : 0
    const minute = date[6] ? Number(date[6]) : 0
    const second = date[8] ? Number(date[8]) : 0
    return new Date(year, month - 1, day, hour, minute, second)
  }
  return content
}

interface ParsedRange {
  sheet: ExcelJS.Worksheet
  startCol: number
  startRow: number
  endCol: number
  endRow: number
}

function parseRange(workbook: ExcelJS.Workbook, range: string): ParsedRange {
  const bang = range.lastIndexOf('!')
  if (bang < 0) throw new Error(`range requires a sheet: ${range}`)
  const rawSheet = range.slice(0, bang)
  const body = range.slice(bang + 1)
  const match = RANGE_LINE.exec(body)
  if (!match) throw new Error(`invalid range: ${range}`)
  const sheet = findSheet(workbook, rawSheet)
  if (!sheet) throw new Error(`sheet not found: ${rawSheet}`)
  return {
    sheet,
    startCol: columnToNumber(match[1]!),
    startRow: Number(match[2]!),
    endCol: columnToNumber(match[3]!),
    endRow: Number(match[4]!),
  }
}

/**
 * Shift selected reference points of a formula. rowDelta/colDelta apply to
 * relative rows/columns; rowThreshold/colThreshold gate the shift so row edits
 * only move references at or below the insertion/deletion point. When
 * editedSheet is set, only references pointing into that sheet are shifted.
 */
export function shiftFormulaReferences(
  formula: string,
  baseSheet: string,
  editedSheet: string | null,
  options: {
    rowDelta?: number
    colDelta?: number
    rowThreshold?: number
    colThreshold?: number
    rowDeletedStart?: number
    rowDeletedEnd?: number
    colDeletedStart?: number
    colDeletedEnd?: number
  } = {},
): string {
  const { rowDelta, colDelta, rowThreshold, colThreshold, rowDeletedStart, rowDeletedEnd, colDeletedStart, colDeletedEnd } = options
  if (
    (rowDelta ?? 0) === 0 &&
    (colDelta ?? 0) === 0 &&
    rowDeletedStart === undefined &&
    colDeletedStart === undefined
  ) return formula
  const hasEquals = formula.trimStart().startsWith('=')
  const raw = hasEquals ? formula.trimStart().slice(1) : formula
  const parsed = parseFormula(`=${raw}`)
  const edits: Array<{ start: number; end: number; text: string }> = []
  for (const ref of parsed.references) {
    const text = raw.slice(ref.range.start, ref.range.end)
    const colon = text.indexOf(':')
    const startToken = colon >= 0 ? text.slice(0, colon) : text
    const endToken = colon >= 0 ? text.slice(colon + 1) : null
    const newStart = shiftPointToken(
      startToken,
      ref.start,
      baseSheet,
      editedSheet,
      { rowDelta, colDelta, rowThreshold, colThreshold, rowDeletedStart, rowDeletedEnd, colDeletedStart, colDeletedEnd },
    )
    if (endToken === null) {
      edits.push({ start: ref.range.start, end: ref.range.end, text: newStart })
    } else {
      const newEnd = shiftPointToken(
        endToken,
        ref.end!,
        baseSheet,
        editedSheet,
        { rowDelta, colDelta, rowThreshold, colThreshold, rowDeletedStart, rowDeletedEnd, colDeletedStart, colDeletedEnd },
      )
      edits.push({ start: ref.range.start, end: ref.range.end, text: `${newStart}:${newEnd}` })
    }
  }
  let result = raw
  for (const edit of [...edits].sort((a, b) => b.start - a.start)) {
    result = `${result.slice(0, edit.start)}${edit.text}${result.slice(edit.end)}`
  }
  return hasEquals ? `=${result}` : result
}

function shiftPointToken(
  token: string,
  point: RefPoint,
  baseSheet: string,
  editedSheet: string | null,
  options: {
    rowDelta?: number
    colDelta?: number
    rowThreshold?: number
    colThreshold?: number
    rowDeletedStart?: number
    rowDeletedEnd?: number
    colDeletedStart?: number
    colDeletedEnd?: number
  },
): string {
  const effectiveSheet = normalizeSheet(point.sheet ?? baseSheet)
  if (editedSheet && effectiveSheet !== normalizeSheet(editedSheet)) return token
  const {
    rowDelta,
    colDelta,
    rowThreshold,
    colThreshold,
    rowDeletedStart,
    rowDeletedEnd,
    colDeletedStart,
    colDeletedEnd,
  } = options

  const colMatch = /^(.*?)(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(token)
  const wholeColMatch = /^(.*?)(\$?)([A-Za-z]{1,3})$/.exec(token)
  const hasRow = colMatch !== null
  const prefix = hasRow ? colMatch![1]! : wholeColMatch?.[1] ?? token
  const absCol = hasRow ? colMatch![2]! === '$' : wholeColMatch?.[2] === '$'
  const col = hasRow ? colMatch![3]! : wholeColMatch?.[3] ?? null
  const absRow = hasRow ? colMatch![4]! === '$' : false
  const row = hasRow ? Number(colMatch![5]!) : null

  const currentColNumber = col ? columnToNumber(col) : null
  if (
    (rowDeletedStart !== undefined && row !== null && !absRow && row >= rowDeletedStart && row <= (rowDeletedEnd ?? rowDeletedStart)) ||
    (colDeletedStart !== undefined && currentColNumber !== null && !absCol && currentColNumber >= colDeletedStart && currentColNumber <= (colDeletedEnd ?? colDeletedStart))
  ) {
    return '#REF!'
  }

  let newCol = col
  if (col && !absCol && colDelta) {
    const current = columnToNumber(col)
    if ((colThreshold === undefined || current >= colThreshold) && current + colDelta >= 1) {
      newCol = numberToColumn(current + colDelta)
    }
  }
  let newRow = row
  if (row !== null && !absRow && rowDelta) {
    if ((rowThreshold === undefined || row >= rowThreshold) && row + rowDelta >= 1) {
      newRow = row + rowDelta
    }
  }
  if (newCol === col && newRow === row) return token
  const colPart = `${absCol ? '$' : ''}${newCol ?? ''}`
  const rowPart = newRow === null ? '' : `${absRow ? '$' : ''}${newRow}`
  return `${prefix}${colPart}${rowPart}`
}

function cellContentOf(cell: ExcelJS.Cell): string {
  if (cell.formula) return `=${cell.formula}`
  const value = cell.value
  if (value === null || value === undefined) return ''
  return typeof value === 'object' ? JSON.stringify(value) : String(value)
}

export async function applyOperationsToWorkbook(
  inputPath: string,
  operations: ExcelOperation[],
  outputPath: string,
): Promise<ApplyOperationsResult> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(inputPath)
  const warnings: OperationWarning[] = []

  operations.forEach((operation, index) => {
    switch (operation.op) {
      case 'set': {
        for (const [id, content] of Object.entries(operation.cells)) {
          writeContent(resolveCell(workbook, id), content)
        }
        break
      }
      case 'fill': {
        applyFill(workbook, operation.source, operation.target)
        break
      }
      case 'insertRows': {
        const sheet = findSheet(workbook, operation.sheet)
        if (!sheet) throw new Error(`sheet not found: ${operation.sheet}`)
        if (operation.row < 1 || operation.count < 1) throw new Error(`invalid insertRows: row=${operation.row} count=${operation.count}`)
        sheet.spliceRows(operation.row, 0, ...Array.from({ length: operation.count }, () => []))
        shiftWorkbookRows(workbook, sheet.name, operation.row, operation.count)
        break
      }
      case 'deleteRows': {
        const sheet = findSheet(workbook, operation.sheet)
        if (!sheet) throw new Error(`sheet not found: ${operation.sheet}`)
        if (operation.row < 1 || operation.count < 1) throw new Error(`invalid deleteRows: row=${operation.row} count=${operation.count}`)
        const end = operation.row + operation.count - 1
        for (const formulaCell of collectDeletedRangeRefs(workbook, sheet.name, operation.row, end)) {
          warnings.push({ op: index, message: `formula ${formulaCell} references a deleted row in ${sheet.name}` })
        }
        markDeletedRowRefs(workbook, sheet.name, operation.row, end)
        sheet.spliceRows(operation.row, operation.count)
        shiftWorkbookRows(workbook, sheet.name, end + 1, -operation.count)
        break
      }
      case 'insertColumns': {
        const sheet = findSheet(workbook, operation.sheet)
        if (!sheet) throw new Error(`sheet not found: ${operation.sheet}`)
        const columnNumber = columnToNumber(operation.column)
        if (columnNumber < 1 || operation.count < 1) throw new Error(`invalid insertColumns: column=${operation.column} count=${operation.count}`)
        sheet.spliceColumns(columnNumber, 0, ...Array.from({ length: operation.count }, () => []))
        shiftWorkbookColumns(workbook, sheet.name, columnNumber, operation.count)
        break
      }
      case 'deleteColumns': {
        const sheet = findSheet(workbook, operation.sheet)
        if (!sheet) throw new Error(`sheet not found: ${operation.sheet}`)
        const columnNumber = columnToNumber(operation.column)
        if (columnNumber < 1 || operation.count < 1) throw new Error(`invalid deleteColumns: column=${operation.column} count=${operation.count}`)
        const end = columnNumber + operation.count - 1
        for (const formulaCell of collectDeletedColumnRefs(workbook, sheet.name, columnNumber, end)) {
          warnings.push({ op: index, message: `formula ${formulaCell} references a deleted column in ${sheet.name}` })
        }
        markDeletedColumnRefs(workbook, sheet.name, columnNumber, end)
        sheet.spliceColumns(columnNumber, operation.count)
        shiftWorkbookColumns(workbook, sheet.name, end + 1, -operation.count)
        break
      }
      case 'addSheet': {
        workbook.addWorksheet(operation.name)
        break
      }
      case 'renameSheet': {
        const sheet = findSheet(workbook, operation.oldName)
        if (!sheet) throw new Error(`sheet not found: ${operation.oldName}`)
        sheet.name = operation.newName
        renameSheetReferences(workbook, operation.oldName, operation.newName)
        break
      }
      case 'deleteSheet': {
        const sheet = findSheet(workbook, operation.name)
        if (!sheet) throw new Error(`sheet not found: ${operation.name}`)
        workbook.removeWorksheet(sheet.id)
        break
      }
      case 'clear': {
        for (const id of operation.cells) resolveCell(workbook, id).value = null
        break
      }
      case 'merge': {
        applyMerge(workbook, operation.range, false)
        break
      }
      case 'unmerge': {
        applyMerge(workbook, operation.range, true)
        break
      }
      case 'copyRange': {
        copyRange(workbook, operation.source, operation.target, operation.move ?? false)
        break
      }
      case 'fillSeries': {
        fillSeries(workbook, operation.start, operation.target, operation.step)
        break
      }
      case 'style': {
        applyStyle(workbook, operation.range, operation.style)
        break
      }
      case 'setColumnWidth': {
        const sheet = findSheet(workbook, operation.sheet)
        if (!sheet) throw new Error(`sheet not found: ${operation.sheet}`)
        sheet.getColumn(operation.column).width = operation.width
        break
      }
      case 'setRowHeight': {
        const sheet = findSheet(workbook, operation.sheet)
        if (!sheet) throw new Error(`sheet not found: ${operation.sheet}`)
        sheet.getRow(operation.row).height = operation.height
        break
      }
      case 'freezePanes': {
        const sheet = findSheet(workbook, operation.sheet)
        if (!sheet) throw new Error(`sheet not found: ${operation.sheet}`)
        const columnNumber = columnToNumber(operation.column)
        sheet.views = [{
          state: 'frozen',
          xSplit: Math.max(0, columnNumber - 1),
          ySplit: Math.max(0, operation.row - 1),
          topLeftCell: `${numberToColumn(columnNumber)}${operation.row}`,
        }]
        break
      }
      case 'findReplace': {
        const count = findReplace(workbook, operation.find, operation.replace, operation.sheet, operation.matchCase ?? false)
        warnings.push({ op: index, message: `findReplace replaced ${count} occurrence(s)` })
        break
      }
      case 'duplicateSheet': {
        duplicateSheet(workbook, operation.name, operation.newName)
        break
      }
      case 'hideSheet': {
        const sheet = findSheet(workbook, operation.name)
        if (!sheet) throw new Error(`sheet not found: ${operation.name}`)
        sheet.state = operation.hidden === false ? 'visible' : 'hidden'
        break
      }
      case 'setTabColor': {
        const sheet = findSheet(workbook, operation.name)
        if (!sheet) throw new Error(`sheet not found: ${operation.name}`)
        sheet.properties.tabColor = { argb: normalizeColor(operation.color) }
        break
      }
      case 'sortRange': {
        sortRange(workbook, operation.range, operation.keys, operation.headerRows ?? 0)
        warnings.push({ op: index, message: 'sortRange moved cell content; formulas outside the range still point to their original addresses' })
        break
      }
      case 'dataValidation': {
        applyDataValidation(workbook, operation)
        break
      }
      case 'conditionalFormatting': {
        applyConditionalFormatting(workbook, operation.range, operation.rules)
        break
      }
      case 'autoFilter': {
        const parsed = parseRange(workbook, operation.range)
        parsed.sheet.autoFilter = {
          from: { row: parsed.startRow, column: parsed.startCol },
          to: { row: parsed.endRow, column: parsed.endCol },
        }
        break
      }
      case 'subtotal': {
        applySubtotal(workbook, operation)
        warnings.push({ op: index, message: 'subtotal groups data by the group column; sort the range by that column first for correct grouping' })
        break
      }
      case 'aggregateReport': {
        applyAggregateReport(workbook, operation)
        break
      }
      case 'filterToRange': {
        applyFilterToRange(workbook, operation)
        break
      }
      case 'protectSheet': {
        const sheet = findSheet(workbook, operation.sheet)
        if (!sheet) throw new Error(`sheet not found: ${operation.sheet}`)
        sheet.protect(operation.password ?? '', { selectLockedCells: true, selectUnlockedCells: true })
        break
      }
      case 'unprotectSheet': {
        const sheet = findSheet(workbook, operation.sheet)
        if (!sheet) throw new Error(`sheet not found: ${operation.sheet}`)
        sheet.unprotect()
        break
      }
      case 'mailMerge': {
        applyMailMerge(workbook, operation)
        break
      }
      case 'addTable': {
        addTable(workbook, operation)
        break
      }
    }
  })

  await workbook.xlsx.writeFile(outputPath)
  return { warnings }
}

function applyFill(workbook: ExcelJS.Workbook, sourceId: string, targetRange: string): void {
  const source = resolveCell(workbook, sourceId)
  const sourceCell = parseCellId(sourceId)
  const bang = targetRange.lastIndexOf('!')
  const rawSheet = bang >= 0 ? targetRange.slice(0, bang) : null
  const body = bang >= 0 ? targetRange.slice(bang + 1) : targetRange
  const match = RANGE_LINE.exec(body)
  if (!match) throw new Error(`invalid fill target: ${targetRange}`)
  const targetSheetName = rawSheet ?? sourceCell.sheet
  const sheet = findSheet(workbook, targetSheetName)
  if (!sheet) throw new Error(`sheet not found: ${targetSheetName}`)
  const startCol = columnToNumber(match[1]!)
  const endCol = columnToNumber(match[3]!)
  const startRow = Number(match[2]!)
  const endRow = Number(match[4]!)
  const content = cellContentOf(source)
  if (!content) return
  for (let col = startCol; col <= endCol; col++) {
    for (let row = startRow; row <= endRow; row++) {
      if (col === columnToNumber(sourceCell.column) && row === sourceCell.row) continue
      const cell = sheet.getCell(`${numberToColumn(col)}${row}`)
      const rowDelta = row - sourceCell.row
      const colDelta = col - columnToNumber(sourceCell.column)
      const value = content.startsWith('=')
        ? shiftFormulaReferences(content, sourceCell.sheet, null, { rowDelta, colDelta })
        : content
      writeContent(cell, value)
    }
  }
}

function shiftWorkbookRows(workbook: ExcelJS.Workbook, editedSheet: string, threshold: number, rowDelta: number): void {
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

function collectDeletedRangeRefs(
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

function shiftWorkbookColumns(workbook: ExcelJS.Workbook, editedSheet: string, threshold: number, colDelta: number): void {
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

function collectDeletedColumnRefs(
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

function markDeletedRowRefs(workbook: ExcelJS.Workbook, editedSheet: string, start: number, end: number): void {
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

function markDeletedColumnRefs(workbook: ExcelJS.Workbook, editedSheet: string, start: number, end: number): void {
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

function sortRange(
  workbook: ExcelJS.Workbook,
  range: string,
  keys: Array<{ column: string; direction?: 'asc' | 'desc' }>,
  headerRows: number,
): void {
  const parsed = parseRange(workbook, range)
  const keyColumns = keys.map((key) => ({
    column: columnToNumber(key.column),
    direction: key.direction ?? 'asc',
  }))
  for (const key of keyColumns) {
    if (key.column < parsed.startCol || key.column > parsed.endCol) {
      throw new Error(`sort key column outside range: ${numberToColumn(key.column)}`)
    }
  }
  if (headerRows < 0 || headerRows >= parsed.endRow - parsed.startRow + 1) {
    throw new Error(`invalid headerRows: ${headerRows}`)
  }
  interface SortRow {
    cells: Record<string, ExcelJS.CellValue>
    keys: Array<string | number | Date | null>
  }
  const rows: SortRow[] = []
  for (let row = parsed.startRow + headerRows; row <= parsed.endRow; row++) {
    const cells: Record<string, ExcelJS.CellValue> = {}
    const keyValues: SortRow['keys'] = []
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const letter = numberToColumn(col)
      const cell = parsed.sheet.getCell(`${letter}${row}`)
      cells[letter] = cell.value
      if (keyColumns.some((key) => key.column === col)) keyValues.push(cell.value as SortRow['keys'][number])
    }
    rows.push({ cells, keys: keyValues })
  }
  rows.sort((a, b) => compareSortRows(a.keys, b.keys, keyColumns.map((key) => key.direction)))
  for (let index = 0; index < rows.length; index++) {
    const targetRow = parsed.startRow + headerRows + index
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const letter = numberToColumn(col)
      parsed.sheet.getCell(`${letter}${targetRow}`).value = rows[index]!.cells[letter] ?? null
    }
  }
}

function compareSortRows(
  a: Array<string | number | Date | null>,
  b: Array<string | number | Date | null>,
  directions: Array<'asc' | 'desc'>,
): number {
  for (let index = 0; index < a.length; index++) {
    const comparison = compareSortValue(a[index]!, b[index]!)
    if (comparison !== 0) return directions[index] === 'desc' ? -comparison : comparison
  }
  return 0
}

function compareSortValue(a: string | number | Date | null, b: string | number | Date | null): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (a instanceof Date && b instanceof Date) return a.getTime() - b.getTime()
  const left = a === null || a === undefined ? '' : String(a)
  const right = b === null || b === undefined ? '' : String(b)
  return left < right ? -1 : left > right ? 1 : 0
}

function applyDataValidation(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'dataValidation' }>,
): void {
  const parsed = parseRange(workbook, options.range)
  const validation: ExcelJS.DataValidation = {
    type: options.type,
    operator: options.operator,
    formulae: [],
    allowBlank: options.allowBlank,
    showInputMessage: options.showInputMessage,
    prompt: options.prompt,
    showErrorMessage: options.showErrorMessage,
    errorStyle: options.errorStyle,
    errorTitle: options.errorTitle,
    error: options.error,
  }
  if (options.type === 'list') {
    if (!options.formula1) throw new Error('list data validation requires formula1 (comma-separated items or a range)')
    validation.formulae = [looksLikeRange(options.formula1) ? options.formula1 : `"${options.formula1}"`]
  } else if (options.formula1 !== undefined) {
    validation.formulae = [options.formula1]
    if (options.formula2 !== undefined) validation.formulae.push(options.formula2)
  }
  for (let row = parsed.startRow; row <= parsed.endRow; row++) {
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      parsed.sheet.getCell(`${numberToColumn(col)}${row}`).dataValidation = validation
    }
  }
}

function looksLikeRange(value: string): boolean {
  return /^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$/.test(value) || /[!$]/.test(value)
}

function applyConditionalFormatting(
  workbook: ExcelJS.Workbook,
  range: string,
  rules: Extract<ExcelOperation, { op: 'conditionalFormatting' }>['rules'],
): void {
  const parsed = parseRange(workbook, range)
  const mapped = rules.map((rule) => {
    const style = rule.style ? excelStyleToWorkbookStyle(rule.style) : undefined
    if (rule.type === 'cellIs') {
      if (!rule.operator || rule.formula === undefined) {
        throw new Error('cellIs conditional formatting requires operator and formula')
      }
      return {
        type: 'cellIs',
        operator: rule.operator,
        formulae: [rule.formula, ...(rule.formula2 !== undefined ? [rule.formula2] : [])],
        style,
      }
    }
    return { type: 'expression', formulae: [String(rule.formula ?? '')], style }
  })
  const ref = `${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${parsed.endRow}`
  parsed.sheet.addConditionalFormatting({ ref, rules: mapped as ExcelJS.ConditionalFormattingRule[] })
}

function excelStyleToWorkbookStyle(style: ExcelStyle): ExcelJS.Style {
  const result: Partial<ExcelJS.Style> = {}
  if (style.bold !== undefined || style.italic !== undefined || style.underline !== undefined || style.fontColor !== undefined) {
    result.font = {
      bold: style.bold,
      italic: style.italic,
      underline: style.underline,
      color: style.fontColor ? { argb: normalizeColor(style.fontColor) } : undefined,
    }
  }
  if (style.fill !== undefined) {
    result.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: normalizeColor(style.fill) } }
  }
  return result as ExcelJS.Style
}

function addTable(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'addTable' }>,
): void {
  const parsed = parseRange(workbook, options.range)
  const ref = `${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${parsed.endRow}`
  const headerRow = options.headerRow ?? true
  const header = headerRow ? parsed.startRow : null
  const dataStart = headerRow ? parsed.startRow + 1 : parsed.startRow
  const columns: Array<{ name: string }> = []
  for (let col = parsed.startCol; col <= parsed.endCol; col++) {
    const letter = numberToColumn(col)
    const nameCell = header ? parsed.sheet.getCell(`${letter}${header}`).value : null
    columns.push({ name: nameCell === null || nameCell === undefined ? `Column${letter}` : String(nameCell) })
  }
  const rows: Array<Array<ExcelJS.CellValue>> = []
  for (let row = dataStart; row <= parsed.endRow; row++) {
    const values: ExcelJS.CellValue[] = []
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      values.push(parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value)
    }
    rows.push(values)
  }
  parsed.sheet.addTable({
    name: options.name,
    ref,
    headerRow,
    totalsRow: options.totalsRow ?? false,
    columns,
    rows,
    style: {
      showRowStripes: options.showRowStripes ?? true,
      showColumnStripes: options.showColumnStripes ?? false,
    },
  })
}

const SUBTOTAL_CODES: Record<string, number> = {
  sum: 9,
  average: 1,
  count: 2,
  max: 4,
  min: 5,
}

function applySubtotal(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'subtotal' }>,
): void {
  const parsed = parseRange(workbook, options.range)
  const groupCol = columnToNumber(options.groupColumn)
  if (groupCol < parsed.startCol || groupCol > parsed.endCol) {
    throw new Error(`subtotal group column outside range: ${options.groupColumn}`)
  }
  for (const summary of options.summaryColumns) {
    const col = columnToNumber(summary.column)
    if (col < parsed.startCol || col > parsed.endCol) {
      throw new Error(`subtotal summary column outside range: ${summary.column}`)
    }
    if (!SUBTOTAL_CODES[summary.function]) throw new Error(`unsupported subtotal function: ${summary.function}`)
  }
  const sheet = parsed.sheet
  const header = parsed.startRow
  const firstData = parsed.startRow + 1
  const lastData = parsed.endRow

  interface SubtotalGroup {
    value: string
    startRow: number
    endRow: number
  }
  const groups: SubtotalGroup[] = []
  let current: SubtotalGroup | null = null
  for (let row = firstData; row <= lastData; row++) {
    const raw = sheet.getCell(`${numberToColumn(groupCol)}${row}`).value
    const key = raw === null || raw === undefined ? '' : String(raw)
    if (!current || current.value !== key) {
      current = { value: key, startRow: row, endRow: row }
      groups.push(current)
    } else {
      current.endRow = row
    }
  }

  for (let groupIndex = 0; groupIndex < groups.length; groupIndex++) {
    const group = groups[groupIndex]!
    const finalStartRow = group.startRow + groupIndex
    const finalEndRow = group.endRow + groupIndex
    const insertRow = finalEndRow + 1
    sheet.spliceRows(insertRow, 0, [])
    const label = sheet.getCell(`${numberToColumn(groupCol)}${insertRow}`)
    label.value = `${group.value} 汇总`
    label.font = { bold: true }
    for (const summary of options.summaryColumns) {
      const col = columnToNumber(summary.column)
      const cell = sheet.getCell(`${numberToColumn(col)}${insertRow}`)
      cell.value = {
        formula: `SUBTOTAL(${SUBTOTAL_CODES[summary.function]},${numberToColumn(col)}${finalStartRow}:${numberToColumn(col)}${finalEndRow})`,
      }
      cell.font = { bold: true }
    }
  }

  if (options.addGrandTotal ?? true) {
    const totalRow = parsed.endRow + groups.length + 1
    sheet.spliceRows(totalRow, 0, [])
    const label = sheet.getCell(`${numberToColumn(groupCol)}${totalRow}`)
    label.value = '总计'
    label.font = { bold: true }
    for (const summary of options.summaryColumns) {
      const col = columnToNumber(summary.column)
      const cell = sheet.getCell(`${numberToColumn(col)}${totalRow}`)
      cell.value = {
        formula: `SUBTOTAL(${SUBTOTAL_CODES[summary.function]},${numberToColumn(col)}${firstData}:${numberToColumn(col)}${lastData + groups.length})`,
      }
      cell.font = { bold: true }
    }
  }
  void header
}

const REPORT_FUNCTIONS: Record<string, string> = {
  sum: 'SUMIFS',
  average: 'AVERAGEIFS',
  count: 'COUNTIFS',
  counta: 'COUNTIFS',
  max: 'MAXIFS',
  min: 'MINIFS',
}

function applyAggregateReport(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'aggregateReport' }>,
): void {
  const parsed = parseRange(workbook, options.source)
  const groupCol = columnToNumber(options.groupColumn)
  const sourceSheet = parsed.sheet.name
  const firstData = parsed.startRow + 1
  const lastData = parsed.endRow
  const groupRange = `${sourceSheet}!$${numberToColumn(groupCol)}$${firstData}:$${numberToColumn(groupCol)}$${lastData}`

  const groupValues: string[] = []
  const seen = new Set<string>()
  for (let row = firstData; row <= lastData; row++) {
    const raw = parsed.sheet.getCell(`${numberToColumn(groupCol)}${row}`).value
    const key = raw === null || raw === undefined ? '' : String(raw)
    if (!seen.has(key)) {
      seen.add(key)
      groupValues.push(key)
    }
  }

  const outputSheetName = options.outputSheet ?? `${sourceSheet}-汇总`
  let output = findSheet(workbook, outputSheetName)
  if (!output) output = workbook.addWorksheet(outputSheetName)
  const groupHeader = String(parsed.sheet.getCell(`${numberToColumn(groupCol)}${parsed.startRow}`).value ?? options.groupColumn)
  output.getCell('A1').value = groupHeader
  output.getCell('A1').font = { bold: true }
  const metricLabels: Record<string, string> = {
    sum: '合计',
    average: '平均',
    count: '计数',
    counta: '非空计数',
    max: '最大',
    min: '最小',
  }
  options.metrics.forEach((metric, index) => {
    const metricCol = columnToNumber(metric.column)
    const header = String(parsed.sheet.getCell(`${numberToColumn(metricCol)}${parsed.startRow}`).value ?? metric.column)
    const cell = output.getCell(`${numberToColumn(2 + index)}1`)
    cell.value = `${header} ${metricLabels[metric.function]}`
    cell.font = { bold: true }
    void metricCol
  })

  for (let index = 0; index < groupValues.length; index++) {
    const row = 2 + index
    const groupCell = output.getCell(`A${row}`)
    groupCell.value = groupValues[index]
    options.metrics.forEach((metric, metricIndex) => {
      const metricCol = columnToNumber(metric.column)
      const metricRange = `${sourceSheet}!$${numberToColumn(metricCol)}$${firstData}:$${numberToColumn(metricCol)}$${lastData}`
      const fn = REPORT_FUNCTIONS[metric.function]!
      const criteria = `A${row}`
      output.getCell(`${numberToColumn(2 + metricIndex)}${row}`).value = {
        formula: `${fn}(${metricRange},${groupRange},${criteria})`,
      }
    })
  }
  const lastGroupRow = 1 + groupValues.length
  options.metrics.forEach((metric, metricIndex) => {
    const col = numberToColumn(2 + metricIndex)
    output.getCell(`${col}${lastGroupRow + 1}`).value = {
      formula: `SUM(${col}2:${col}${lastGroupRow})`,
    }
    output.getCell(`${col}${lastGroupRow + 1}`).font = { bold: true }
  })
  output.getCell(`A${lastGroupRow + 1}`).value = '总计'
  output.getCell(`A${lastGroupRow + 1}`).font = { bold: true }
}

function applyFilterToRange(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'filterToRange' }>,
): void {
  const parsed = parseRange(workbook, options.source)
  const target = parseTargetCell(workbook, options.target, parsed.sheet.name)
  const matchAll = options.matchAll ?? true

  const headerRow: Array<ExcelJS.CellValue> = []
  for (let col = parsed.startCol; col <= parsed.endCol; col++) {
    headerRow.push(parsed.sheet.getCell(`${numberToColumn(col)}${parsed.startRow}`).value)
  }
  let targetRow = target.row
  headerRow.forEach((value, index) => {
    target.sheet.getCell(`${numberToColumn(target.col + index)}${targetRow}`).value = value
  })
  targetRow += 1

  for (let row = parsed.startRow + 1; row <= parsed.endRow; row++) {
    let matched = matchAll
    for (const criterion of options.criteria) {
      const col = columnToNumber(criterion.column)
      const actual = parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value
      const ok = matchesCriterion(actual, criterion.operator, criterion.value)
      if (matchAll && !ok) {
        matched = false
        break
      }
      if (!matchAll && ok) {
        matched = true
        break
      }
    }
    if (!matched) continue
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      target.sheet.getCell(`${numberToColumn(target.col + (col - parsed.startCol))}${targetRow}`).value =
        parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value
    }
    targetRow += 1
  }
}

function matchesCriterion(
  actual: ExcelJS.CellValue,
  operator: 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains',
  expected: string | number,
): boolean {
  const actualNumber = typeof actual === 'number' ? actual : null
  const expectedNumber = typeof expected === 'number' ? expected : Number(expected)
  const actualText = actual === null || actual === undefined ? '' : String(actual)
  const expectedText = String(expected)
  switch (operator) {
    case 'eq':
      return actualNumber !== null && Number.isFinite(expectedNumber)
        ? actualNumber === expectedNumber
        : actualText.toLowerCase() === expectedText.toLowerCase()
    case 'neq':
      return !matchesCriterion(actual, 'eq', expected)
    case 'contains':
      return actualText.toLowerCase().includes(expectedText.toLowerCase())
    case 'gt':
      return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber > expectedNumber
    case 'gte':
      return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber >= expectedNumber
    case 'lt':
      return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber < expectedNumber
    case 'lte':
      return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber <= expectedNumber
  }
}

function applyMailMerge(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'mailMerge' }>,
): void {
  const template = parseRange(workbook, options.template)
  const data = parseRange(workbook, options.data)
  const headers = new Map<string, number>()
  for (let col = data.startCol; col <= data.endCol; col++) {
    const raw = data.sheet.getCell(`${numberToColumn(col)}${data.startRow}`).value
    headers.set(String(raw ?? '').toLowerCase(), col)
  }
  const templateRows: Array<Record<number, ExcelJS.CellValue>> = []
  for (let row = template.startRow; row <= template.endRow; row++) {
    const cells: Record<number, ExcelJS.CellValue> = {}
    for (let col = template.startCol; col <= template.endCol; col++) {
      cells[col] = template.sheet.getCell(`${numberToColumn(col)}${row}`).value
    }
    templateRows.push(cells)
  }

  const outputSheetName = options.outputSheet ?? `${template.sheet.name}-合并`
  let output = findSheet(workbook, outputSheetName)
  if (!output) output = workbook.addWorksheet(outputSheetName)
  let outputRow = 1
  const placeholder = /\{([^{}]+)\}/g
  for (let dataRow = data.startRow + 1; dataRow <= data.endRow; dataRow++) {
    const record = new Map<string, ExcelJS.CellValue>()
    for (const [header, col] of headers) {
      record.set(header, data.sheet.getCell(`${numberToColumn(col)}${dataRow}`).value)
    }
    for (const templateRow of templateRows) {
      for (const [col, value] of Object.entries(templateRow)) {
        const column = Number(col)
        const text = value === null || value === undefined ? '' : String(value)
        if (/^\{[^{}]+\}$/.test(text.trim())) {
          const key = text.trim().slice(1, -1).toLowerCase()
          output.getCell(`${numberToColumn(column)}${outputRow}`).value = record.get(key) ?? text
          continue
        }
        if (placeholder.test(text)) {
          placeholder.lastIndex = 0
          output.getCell(`${numberToColumn(column)}${outputRow}`).value = text.replace(placeholder, (_match, name: string) => {
            const replacement = record.get(String(name).toLowerCase())
            return replacement === undefined ? _match : String(replacement)
          })
          continue
        }
        output.getCell(`${numberToColumn(column)}${outputRow}`).value = value
      }
      outputRow += 1
    }
  }
}

function parseTargetCell(
  workbook: ExcelJS.Workbook,
  target: string,
  defaultSheet: string,
): { sheet: ExcelJS.Worksheet; col: number; row: number } {
  const bang = target.lastIndexOf('!')
  const sheetName = bang >= 0 ? target.slice(0, bang) : defaultSheet
  const body = bang >= 0 ? target.slice(bang + 1) : target
  const match = /^([A-Za-z]{1,3})(\d+)$/.exec(body)
  if (!match) throw new Error(`invalid target cell: ${target}`)
  const sheet = findSheet(workbook, sheetName)
  if (!sheet) throw new Error(`sheet not found: ${sheetName}`)
  return { sheet, col: columnToNumber(match[1]!), row: Number(match[2]!) }
}

function copyRange(workbook: ExcelJS.Workbook, sourceRange: string, targetCell: string, move: boolean): void {
  const parsed = parseRange(workbook, sourceRange)
  const bang = targetCell.lastIndexOf('!')
  const targetSheetName = bang >= 0 ? targetCell.slice(0, bang) : parsed.sheet.name
  const targetBody = bang >= 0 ? targetCell.slice(bang + 1) : targetCell
  const match = /^([A-Za-z]{1,3})(\d+)$/.exec(targetBody)
  if (!match) throw new Error(`invalid target cell: ${targetCell}`)
  const targetSheet = findSheet(workbook, targetSheetName)
  if (!targetSheet) throw new Error(`sheet not found: ${targetSheetName}`)
  const targetCol = columnToNumber(match[1]!)
  const targetRow = Number(match[2]!)

  for (let row = parsed.startRow; row <= parsed.endRow; row++) {
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const source = parsed.sheet.getCell(`${numberToColumn(col)}${row}`)
      const destCol = targetCol + (col - parsed.startCol)
      const destRow = targetRow + (row - parsed.startRow)
      const dest = targetSheet.getCell(`${numberToColumn(destCol)}${destRow}`)
      const content = cellContentOf(source)
      if (!content) {
        dest.value = null
        continue
      }
      dest.value = content.startsWith('=')
        ? {
            formula: shiftFormulaReferences(content, parsed.sheet.name, null, {
              rowDelta: destRow - row,
              colDelta: destCol - col,
            }).slice(1),
          }
        : toCellValue(content)
    }
  }
  if (move) {
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
      for (let col = parsed.startCol; col <= parsed.endCol; col++) {
        parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value = null
      }
    }
  }
}

function fillSeries(workbook: ExcelJS.Workbook, startId: string, targetRange: string, step?: number): void {
  const startCell = resolveCell(workbook, startId)
  const startParsed = parseCellId(startId)
  const range = parseRange(workbook, targetRange)
  const startCol = columnToNumber(startParsed.column)
  if (startParsed.row !== range.startRow || startCol !== range.startCol) {
    throw new Error('fillSeries start cell must be the top-left cell of the target range')
  }
  const base = typeof startCell.value === 'number'
    ? startCell.value
    : startCell.value instanceof Date
      ? startCell.value.getTime()
      : null
  if (base === null) throw new Error('fillSeries start cell must be a number or date')
  const isDate = startCell.value instanceof Date
  const stepValue = step ?? (isDate ? 86_400_000 : 1)
  let index = 0
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      if (row === startParsed.row && col === startCol) continue
      index += 1
      range.sheet.getCell(`${numberToColumn(col)}${row}`).value = isDate
        ? new Date(base + stepValue * index)
        : base + stepValue * index
    }
  }
}

function applyStyle(workbook: ExcelJS.Workbook, range: string, style: ExcelStyle): void {
  const parsed = parseRange(workbook, range)
  for (let row = parsed.startRow; row <= parsed.endRow; row++) {
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`)
      const font = cell.font ?? {}
      if (style.bold !== undefined || style.italic !== undefined || style.underline !== undefined || style.fontColor !== undefined) {
        cell.font = {
          ...font,
          bold: style.bold ?? font.bold,
          italic: style.italic ?? font.italic,
          underline: style.underline ?? font.underline,
          color: style.fontColor ? { argb: normalizeColor(style.fontColor) } : font.color,
        }
      }
      if (style.fill !== undefined) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: normalizeColor(style.fill) } }
      }
      if (style.numberFormat !== undefined) cell.numFmt = style.numberFormat
      const alignment = cell.alignment ?? {}
      if (style.hAlign !== undefined || style.vAlign !== undefined || style.wrapText !== undefined) {
        cell.alignment = {
          ...alignment,
          horizontal: style.hAlign ?? alignment.horizontal,
          vertical: style.vAlign ?? alignment.vertical,
          wrapText: style.wrapText ?? alignment.wrapText,
        }
      }
    }
  }
}

function normalizeColor(color: string): string {
  const hex = color.replace('#', '').trim()
  if (/^[0-9A-Fa-f]{6}$/.test(hex)) return `FF${hex.toUpperCase()}`
  if (/^[0-9A-Fa-f]{8}$/.test(hex)) return hex.toUpperCase()
  throw new Error(`invalid color: ${color} (use 6-digit hex like FF0000)`)
}

function findReplace(
  workbook: ExcelJS.Workbook,
  find: string,
  replace: string,
  sheetName: string | undefined,
  matchCase: boolean,
): number {
  let count = 0
  const visit = (sheet: ExcelJS.Worksheet): void => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const content = cellContentOf(cell)
        if (!content) return
        const replaced = replaceAllCase(content, find, replace, matchCase)
        if (replaced === content) return
        count += 1
        cell.value = content.startsWith('=')
          ? { formula: replaced.slice(1) }
          : toCellValue(replaced)
      })
    })
  }
  if (sheetName) {
    const sheet = findSheet(workbook, sheetName)
    if (!sheet) throw new Error(`sheet not found: ${sheetName}`)
    visit(sheet)
  } else {
    workbook.eachSheet(visit)
  }
  return count
}

function replaceAllCase(text: string, find: string, replace: string, matchCase: boolean): string {
  if (matchCase) return text.replaceAll(find, replace)
  const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return text.replace(new RegExp(escaped, 'gi'), replace)
}

function duplicateSheet(workbook: ExcelJS.Workbook, name: string, newName: string): void {
  const source = findSheet(workbook, name)
  if (!source) throw new Error(`sheet not found: ${name}`)
  if (findSheet(workbook, newName)) throw new Error(`sheet already exists: ${newName}`)
  const copy = workbook.addWorksheet(newName)
  source.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      copy.getCell(cell.address).value = cell.value
    })
  })
  for (const merged of source.model.merges ?? []) copy.mergeCells(merged)
}

function renameSheetReferences(workbook: ExcelJS.Workbook, oldName: string, newName: string): void {
  const oldQuoted = `'${oldName.replace(/'/g, "''")}'!`
  const newQuoted = `'${newName.replace(/'/g, "''")}'!`
  const oldBare = `${oldName}!`
  const newBare = `${newName}!`
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!cell.formula) return
        const formula = cell.formula
          .replaceAll(oldQuoted, newQuoted)
          .replaceAll(oldBare, newBare)
        if (formula !== cell.formula) cell.value = { formula }
      })
    })
  })
}

function applyMerge(workbook: ExcelJS.Workbook, range: string, unmerge: boolean): void {
  const bang = range.lastIndexOf('!')
  const rawSheet = bang >= 0 ? range.slice(0, bang) : null
  const body = bang >= 0 ? range.slice(bang + 1) : range
  if (!rawSheet) throw new Error(`merge range requires a sheet: ${range}`)
  const sheet = findSheet(workbook, rawSheet)
  if (!sheet) throw new Error(`sheet not found: ${rawSheet}`)
  if (unmerge) sheet.unMergeCells(body)
  else sheet.mergeCells(body)
}

export async function operateWorkbookFile(
  path: string,
  operations: ExcelOperation[],
  outputPath: string,
): Promise<OperateResult> {
  const result = await applyOperationsToWorkbook(path, operations, outputPath)
  const cells = await readWorkbookCells(await readFile(outputPath))
  const validation = validate(cells)
  return { ...result, outputPath, validation }
}
