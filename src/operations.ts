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
  options: { rowDelta?: number; colDelta?: number; rowThreshold?: number; colThreshold?: number } = {},
): string {
  const { rowDelta, colDelta, rowThreshold, colThreshold } = options
  if ((rowDelta ?? 0) === 0 && (colDelta ?? 0) === 0) return formula
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
      { rowDelta, colDelta, rowThreshold, colThreshold },
    )
    if (endToken === null) {
      edits.push({ start: ref.range.start, end: ref.range.end, text: newStart })
    } else {
      const newEnd = shiftPointToken(
        endToken,
        ref.end!,
        baseSheet,
        editedSheet,
        { rowDelta, colDelta, rowThreshold, colThreshold },
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
  options: { rowDelta?: number; colDelta?: number; rowThreshold?: number; colThreshold?: number },
): string {
  const effectiveSheet = normalizeSheet(point.sheet ?? baseSheet)
  if (editedSheet && effectiveSheet !== normalizeSheet(editedSheet)) return token
  const { rowDelta, colDelta, rowThreshold, colThreshold } = options

  const colMatch = /^(.*?)(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(token)
  const wholeColMatch = /^(.*?)(\$?)([A-Za-z]{1,3})$/.exec(token)
  const hasRow = colMatch !== null
  const prefix = hasRow ? colMatch![1]! : wholeColMatch?.[1] ?? token
  const absCol = hasRow ? colMatch![2]! === '$' : wholeColMatch?.[2] === '$'
  const col = hasRow ? colMatch![3]! : wholeColMatch?.[3] ?? null
  const absRow = hasRow ? colMatch![4]! === '$' : false
  const row = hasRow ? Number(colMatch![5]!) : null

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
