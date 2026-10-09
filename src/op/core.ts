/**
 * Cells, ranges and the small helpers every operation needs.
 *
 * The bottom of the stack: `operations.ts` used to hold all of this alongside the
 * dispatcher and 58 handlers. Nothing here calls anything above it, which is what
 * lets the rest be split without cycles.
 */
import ExcelJS from 'exceljs'
import { columnToNumber, normalizeSheet, numberToColumn, parseCellId, parseFormula, type RefPoint } from '../formula.ts'
import { cellContent, contentToCellValue } from '../workbook.ts'
import { t } from '../i18n.ts'
export const RANGE_LINE = /^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/

export function findSheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined {
  const normalized = normalizeSheet(name)
  return workbook.worksheets.find((sheet) => normalizeSheet(sheet.name) === normalized)
}

export function resolveCell(workbook: ExcelJS.Workbook, id: string): ExcelJS.Cell {
  const parsed = parseCellId(id)
  const sheet = findSheet(workbook, parsed.sheet)
  if (!sheet) throw new Error(`sheet not found: ${parsed.sheet}`)
  return sheet.getCell(`${parsed.column}${parsed.row}`)
}

export function writeContent(cell: ExcelJS.Cell, content: unknown): void {
  // Direct callers (excel_operate "set") may pass typed scalars: the tool schema
  // documents numbers/dates/booleans as first-class content. Only strings carry
  // text to trim and infer, so assign scalars as-is instead of calling .trim().
  if (typeof content !== 'string') {
    cell.value = toScalarValue(content)
    return
  }
  const trimmed = content.trim()
  cell.value = contentToCellValue(trimmed)
}

/**
 * Keep typed scalars intact. Non-finite numbers degrade to text because Excel
 * cannot represent NaN/Infinity and the resulting workbook would not open.
 */
function toScalarValue(content: unknown): ExcelJS.CellValue {
  if (content === null || content === undefined) return null
  if (typeof content === 'number') return Number.isFinite(content) ? content : String(content)
  if (typeof content === 'boolean' || content instanceof Date) return content
  return String(content)
}

/**
 * Convert user-provided text into an Excel value: formulas stay formulas,
 * plain numbers/dates/booleans keep their type, everything else is text.
 * Workplace spreadsheets break when "100" is written as text, so numeric
 * strings are typed before they reach ExcelJS. The conversion itself lives next
 * to `cellContent` in workbook.ts so every writer shares it.
 */

interface ParsedRange {
  sheet: ExcelJS.Worksheet
  startCol: number
  startRow: number
  endCol: number
  endRow: number
}

export function parseRange(workbook: ExcelJS.Workbook, range: string): ParsedRange {
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

export function cellContentOf(cell: ExcelJS.Cell): string {
  if (cell.formula) return `=${cell.formula}`
  const value = cell.value
  if (value === null || value === undefined) return ''
  return typeof value === 'object' ? JSON.stringify(value) : String(value)
}

/**
 * Qualify a sheet name for embedding inside an Excel formula string. Names
 * that are not a plain identifier (spaces, punctuation, leading digits, CJK
 * beyond letters) must be single-quoted with internal quotes doubled, or the
 * generated reference breaks and Excel repairs/drops it.
 */
export function qualifySheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) return name
  return `'${name.replaceAll("'", "''")}'`
}

/** Build an absolute A1 column range for embedding in a formula: Sheet!$C$2:$C$9. */
export function absoluteColumnRef(sheetName: string, colLetter: string, fromRow: number, toRow: number): string {
  return `${qualifySheetName(sheetName)}!$${colLetter}$${fromRow}:$${colLetter}$${toRow}`
}

/** Delete rows with the same reference-shift semantics as the deleteRows op. */
export function properCase(text: string): string {
  return text.toLowerCase().replace(/(^|\s)(\S)/g, (_match, sep: string, char: string) => `${sep}${char.toUpperCase()}`)
}

/** Fullwidth ASCII/space/punctuation to halfwidth, then trim and collapse spaces. */
export function normalizeTextValue(text: string): string {
  let out = ''
  for (const char of text) {
    const code = char.charCodeAt(0)
    if (code >= 0xff01 && code <= 0xff5e) out += String.fromCharCode(code - 0xfee0)
    else if (char === '\u3000') out += ' '
    else if (char === '\u2018' || char === '\u2019') out += "'"
    else if (char === '\u201c' || char === '\u201d') out += '"'
    else out += char
  }
  return out.trim().replace(/\s+/g, ' ')
}

/**
 * Normalized similarity in [0, 1] for fuzzy matching: exact match is 1,
 * otherwise 1 minus the Levenshtein distance ratio over the longer string.
 * Callers normalize (trim/lowercase) before calling.
 */
export function similarity(a: string, b: string): number {
  if (a === b) return 1
  if (a.length === 0 || b.length === 0) return 0
  if (a.length > b.length) return similarity(b, a)
  const previous = Array.from({ length: a.length + 1 }, (_, i) => i)
  let current = new Array<number>(a.length + 1)
  for (let j = 1; j <= b.length; j++) {
    current[0] = j
    for (let i = 1; i <= a.length; i++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1
      current[i] = Math.min(previous[i]! + 1, current[i - 1]! + 1, previous[i - 1]! + cost)
    }
    for (let i = 0; i <= a.length; i++) previous[i] = current[i]!
  }
  return 1 - (previous[a.length]! / b.length)
}

export function parseTargetCell(
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

export function splitByWidth(text: string, widths: number[]): string[] {
  const parts: string[] = []
  let offset = 0
  for (const width of widths) {
    parts.push(text.slice(offset, offset + width).trim())
    offset += width
  }
  if (offset < text.length) parts.push(text.slice(offset).trim())
  return parts
}

/** Turn "Sheet2!A1" / "#明细!B2" into the HYPERLINK target "#'Sheet 2'!A1". */
/** Clone font/fill/border/alignment/number format from one cell onto every cell in the target range. */
/** Replace formulas with their cached results ("paste values" in place). */
/** Write the distinct values of a source column into a target column, first-seen order. */
/**
 * Dedup key that does not collapse distinct types: number 1, text "1" and
 * boolean TRUE get separate keys; a formula with no cached result keys on its
 * formula text instead of collapsing every such cell to "".
 */
/** Patch every existing sheet view without dropping frozen panes or other flags. */
/** Reorder sheets by rewriting orderNo (worksheets getter sorts by it). */
/** Append a live RANK column next to a metric column. */

