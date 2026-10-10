/**
 * Editing cells: filling, copying, merging, transposing, clearing, finding text.
 *
 * These are the operations that move or rewrite content in place, as opposed to the
 * ones that compute something. They share the reference-shifting question — does the
 * edit move other cells with it — which is why they read together. Split out of
 * `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs'
import { columnToNumber, numberToColumn, parseCellId, parseFormula } from '../formula.ts'
import { t } from '../i18n.ts'
import { cellContent, contentToCellValue } from '../workbook.ts'
import { writeWorkbookSafely } from '../safe-write.ts'
import type { ExcelOperation, OperationWarning } from '../operation-types.ts'
import { RANGE_LINE, cellContentOf, findSheet, parseRange, parseTargetCell, qualifySheetName, resolveCell, shiftFormulaReferences, writeContent } from './core.ts'
import { displayWidth } from './layout.ts'

export function applyFill(workbook: ExcelJS.Workbook, sourceId: string, targetRange: string): void {
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
  const startRow = Number(match[2]!)
  // No colon: filling that one cell, which is a 1×1 target.
  const endCol = match[3] ? columnToNumber(match[3]) : startCol
  const endRow = match[4] ? Number(match[4]) : startRow
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

export function copyRange(workbook: ExcelJS.Workbook, sourceRange: string, targetCell: string, move: boolean, valuesOnly = false, warnings?: { push(w: OperationWarning): void }, opIndex = 0): void {
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

  // Snapshot the source block before writing: copying onto an overlapping
  // range (down/right) would otherwise read cells that earlier writes already
  // replaced, corrupting the result.
  const snapshot: Array<{ row: number; col: number; formula?: string; result?: ExcelJS.CellValue; value: ExcelJS.CellValue; content: string }> = []
  for (let row = parsed.startRow; row <= parsed.endRow; row++) {
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const source = parsed.sheet.getCell(`${numberToColumn(col)}${row}`)
      snapshot.push({ row, col, formula: source.formula, result: source.formula ? source.result : undefined, value: source.value, content: cellContentOf(source) })
    }
  }

  for (const { row, col, formula, result, value, content } of snapshot) {
    const destCol = targetCol + (col - parsed.startCol)
    const destRow = targetRow + (row - parsed.startRow)
    const dest = targetSheet.getCell(`${numberToColumn(destCol)}${destRow}`)
    if (valuesOnly) {
      // Paste-special: values only. Formulas contribute their last cached
      // result; empty cells clear the destination.
      if (formula) {
        if (result === undefined || result === null) {
          // No cached value (common for freshly written formulas): fall back
          // to copying the shifted formula so nothing is lost.
          dest.value = {
            formula: shiftFormulaReferences(content, parsed.sheet.name, null, {
              rowDelta: destRow - row,
              colDelta: destCol - col,
            }).slice(1),
          }
          warnings?.push({ op: opIndex, message: t('copyRange valuesOnly：部分公式无缓存结果，已按公式复制') })
        } else {
          dest.value = result
        }
      } else {
        dest.value = value
      }
      continue
    }
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
      : contentToCellValue(content)
  }
  if (move) {
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
      for (let col = parsed.startCol; col <= parsed.endCol; col++) {
        // Clear the source only for cells that the destination did not also
        // write into (fully non-overlapping copies); overlapping in-place moves
        // keep the copied block intact.
        const cleared = targetCol + (col - parsed.startCol)
        const clearedRow = targetRow + (row - parsed.startRow)
        const outsideDest =
          cleared < parsed.startCol || cleared > parsed.endCol ||
          clearedRow < parsed.startRow || clearedRow > parsed.endRow ||
          targetSheet.name !== parsed.sheet.name
        if (outsideDest) parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value = null
      }
    }
  }
}

export function fillSeries(workbook: ExcelJS.Workbook, startId: string, targetRange: string, step?: number): void {
  const startCell = resolveCell(workbook, startId)
  const startParsed = parseCellId(startId)
  const range = parseRange(workbook, targetRange)
  const startCol = columnToNumber(startParsed.column)
  if (startParsed.row !== range.startRow || startCol !== range.startCol) {
    throw new Error('fillSeries start cell must be the top-left cell of the target range')
  }
  const startContent = cellContentOf(startCell)
  if (startContent.startsWith('=')) {
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        if (row === startParsed.row && col === startCol) continue
        const cell = range.sheet.getCell(`${numberToColumn(col)}${row}`)
        const shifted = shiftFormulaReferences(startContent, startParsed.sheet, null, {
          rowDelta: row - startParsed.row,
          colDelta: col - startCol,
        })
        writeContent(cell, shifted)
      }
    }
    return
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

export function findReplace(
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
          : contentToCellValue(replaced)
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

export function duplicateSheet(workbook: ExcelJS.Workbook, name: string, newName: string): void {
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

export function renameSheetReferences(workbook: ExcelJS.Workbook, oldName: string, newName: string): void {
  const oldQuoted = `'${oldName.replace(/'/g, "''")}'!`
  const newQuoted = `'${newName.replace(/'/g, "''")}'!`
  const newBare = `${newName}!`
  // Bare references need token boundaries so renaming "A" does not corrupt
  // "AA!" (oldBare is a substring of "AA!"). Match only when the name is not
  // preceded by an identifier char / quote / $ and is followed by a cell ref.
  const bareRef = new RegExp(`(?<![A-Za-z0-9_$'])${escapeRegExp(oldName)}!(?=[A-Za-z$])`, 'g')
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!cell.formula) return
        const formula = cell.formula
          .replaceAll(oldQuoted, newQuoted)
          .replace(bareRef, () => newBare)
        if (formula !== cell.formula) cell.value = { formula }
      })
    })
  })
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

export function applyMerge(workbook: ExcelJS.Workbook, range: string, unmerge: boolean): void {
  const bang = range.lastIndexOf('!')
  const rawSheet = bang >= 0 ? range.slice(0, bang) : null
  const body = bang >= 0 ? range.slice(bang + 1) : range
  if (!rawSheet) throw new Error(`merge range requires a sheet: ${range}`)
  const sheet = findSheet(workbook, rawSheet)
  if (!sheet) throw new Error(`sheet not found: ${rawSheet}`)
  if (unmerge) sheet.unMergeCells(body)
  else sheet.mergeCells(body)
}


export function transposeRange(workbook: ExcelJS.Workbook, sourceRange: string, targetCell: string): void {
  const parsed = parseRange(workbook, sourceRange)
  const target = parseTargetCell(workbook, targetCell, parsed.sheet.name)
  // Snapshot the whole source block first: transposing onto the source (or an
  // overlapping area) must not read cells that earlier writes already replaced.
  const snapshot: Array<{ row: number; col: number; content: string; raw: ExcelJS.CellValue }> = []
  for (let row = parsed.startRow; row <= parsed.endRow; row++) {
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const source = parsed.sheet.getCell(`${numberToColumn(col)}${row}`)
      snapshot.push({ row, col, content: cellContentOf(source), raw: source.value })
    }
  }
  // Transposing onto the source footprint: clear the source block first so a
  // non-square source (e.g. 2x4 -> 4x2) does not leave stale cells behind.
  if (target.sheet.name === parsed.sheet.name) {
    const destLastRow = target.row + (parsed.endCol - parsed.startCol)
    const destLastCol = target.col + (parsed.endRow - parsed.startRow)
    const overlaps =
      target.row <= parsed.endRow && destLastRow >= parsed.startRow &&
      target.col <= parsed.endCol && destLastCol >= parsed.startCol
    if (overlaps) {
      for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
          parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value = null
        }
      }
    }
  }
  for (const { row, col, content, raw } of snapshot) {
    // (row,col) maps to (targetRow + colOffset, targetCol + rowOffset).
    const destRow = target.row + (col - parsed.startCol)
    const destCol = target.col + (row - parsed.startRow)
    const dest = target.sheet.getCell(`${numberToColumn(destCol)}${destRow}`)
    if (!content) continue
    dest.value = content.startsWith('=')
      ? contentToCellValue(shiftFormulaReferences(content, parsed.sheet.name, null, {
          rowDelta: destRow - row,
          colDelta: destCol - col,
        }))
      : raw
  }
}

export function clearRange(workbook: ExcelJS.Workbook, range: string, mode: 'contents' | 'formats' | 'all'): void {
  const parsed = parseRange(workbook, range)
  for (let row = parsed.startRow; row <= parsed.endRow; row++) {
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`)
      if (mode === 'contents') {
        cell.value = null
      } else if (mode === 'formats') {
        cell.style = {}
      } else {
        cell.value = null
        cell.style = {}
      }
    }
  }
}
