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
  | { op: 'addSheet'; name: string }
  | { op: 'renameSheet'; oldName: string; newName: string }
  | { op: 'deleteSheet'; name: string }
  | { op: 'clear'; cells: string[] }
  | { op: 'merge'; range: string }
  | { op: 'unmerge'; range: string }

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
  cell.value = trimmed.startsWith('=')
    ? { formula: trimmed.slice(1) }
    : trimmed
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
