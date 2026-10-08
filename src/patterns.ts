import { canonicalCellId, columnToNumber, parseCellId, parseFormula, type ParsedFormula, type RefPoint } from './formula.ts'
import { excelErrorPattern } from './error-values.ts'

export type PatternAnomalyKind =
  | 'reference-offset'
  | 'structure-mismatch'
  | 'hardcode-break'
  | 'empty-gap'
  | 'circular-reference'
  | 'error-value'

export interface PatternAnomaly {
  kind: PatternAnomalyKind
  cell: string
  message: string
  expected: string | null
  actual: string | null
  confidence: number | null
  slot?: string
  expectedOffsets?: { colOffset: number | null; rowOffset: number | null }
  actualOffsets?: { colOffset: number | null; rowOffset: number | null }
}

export interface ColumnPatternReport {
  sheet: string
  column: string
  cellCount: number
  expected: Record<string, string>
  anomalies: PatternAnomaly[]
}

interface NormalizedRef {
  sheet: string
  colOffset: number | null
  rowOffset: number | null
}

interface FormulaEntry {
  id: string
  row: number
  slots: Map<string, NormalizedRef>
}

function normalizeFormula(
  parsed: ParsedFormula,
  baseSheet: string,
  baseColumnNumber: number,
  baseRow: number,
): Map<string, NormalizedRef> {
  const slots = new Map<string, NormalizedRef>()
  parsed.references.forEach((ref, refIndex) => {
    const points: Array<['start' | 'end', RefPoint | null]> = [['start', ref.start]]
    if (ref.end) points.push(['end', ref.end])
    for (const [role, point] of points) {
      if (!point) continue
      slots.set(`${refIndex}.${role}`, {
        sheet: (point.sheet ?? baseSheet).toUpperCase(),
        colOffset: point.absColumn ? null : columnToNumber(point.column) - baseColumnNumber,
        rowOffset: point.row === null || point.absRow ? null : point.row - baseRow,
      })
    }
  })
  return slots
}

function normalizedKey(value: NormalizedRef): string {
  return JSON.stringify([value.sheet, value.colOffset, value.rowOffset])
}

function formatNormalized(value: NormalizedRef): string {
  const parts: string[] = []
  if (value.sheet) parts.push(value.sheet)
  parts.push(value.colOffset === null ? 'col:abs' : `col:${formatOffset(value.colOffset)}`)
  parts.push(value.rowOffset === null ? 'row:abs' : `row:${formatOffset(value.rowOffset)}`)
  return parts.join(' ')
}

function formatOffset(offset: number): string {
  if (offset === 0) return '0'
  return offset > 0 ? `+${offset}` : `${offset}`
}

export function detectPatternAnomalies(cells: Record<string, string>): ColumnPatternReport[] {
  // Summary rows (subtotal/grand-total labels) legitimately deviate from the
  // column pattern (e.g. a 总计 row sums the summary rows above it). Skip any
  // formula on a row whose label column contains a summary marker.
  const summaryRows = new Set<string>()
  for (const [id, content] of Object.entries(cells)) {
    const trimmed = content.trim()
    if (!/^(总计|小计)$/.test(trimmed) && !trimmed.endsWith('汇总')) continue
    try {
      const cell = parseCellId(id)
      summaryRows.add(`${cell.sheet}:${cell.row}`)
    } catch {
      // malformed ids are skipped below
    }
  }
  const groups = new Map<string, Array<{ id: string; row: number; parsed: ParsedFormula }>>()
  for (const [id, content] of Object.entries(cells)) {
    const trimmed = content.trim()
    if (!trimmed.startsWith('=')) continue
    // SUBTOTAL rows are aggregate summary rows by design (subtotal/grand total);
    // they follow a different shape than data rows and should not be flagged.
    if (/SUBTOTAL\(/i.test(trimmed)) continue
    let cell
    let parsed
    try {
      cell = parseCellId(id)
      if (summaryRows.has(`${cell.sheet}:${cell.row}`)) continue
      parsed = parseFormula(trimmed)
    } catch {
      continue
    }
    const key = `${cell.sheet}:${cell.column}`
    if (!groups.has(key)) groups.set(key, [])
    groups.get(key)!.push({ id, row: cell.row, parsed })
  }

  const reports: ColumnPatternReport[] = []
  for (const [key, entries] of groups) {
    if (entries.length < 2) continue
    const separator = key.indexOf(':')
    const sheet = key.slice(0, separator)
    const column = key.slice(separator + 1)
    entries.sort((a, b) => a.row - b.row)
    const baseColumnNumber = columnToNumber(column)
    const normalized: FormulaEntry[] = entries.map((entry) => ({
      id: entry.id,
      row: entry.row,
      slots: normalizeFormula(entry.parsed, sheet, baseColumnNumber, entry.row),
    }))

    const slotKeys = new Set<string>()
    for (const entry of normalized) {
      for (const slot of entry.slots.keys()) slotKeys.add(slot)
    }

    const expected: Record<string, string> = {}
    const anomalies: PatternAnomaly[] = []
    for (const slotKey of slotKeys) {
      const present = normalized.filter((entry) => entry.slots.has(slotKey))
      const counts = new Map<string, { count: number; value: NormalizedRef }>()
      for (const entry of present) {
        const value = entry.slots.get(slotKey)!
        const bucketKey = normalizedKey(value)
        const bucket = counts.get(bucketKey)
        if (bucket) bucket.count += 1
        else counts.set(bucketKey, { count: 1, value })
      }
      let majority = { count: -1, value: present[0]!.slots.get(slotKey)! }
      for (const bucket of counts.values()) {
        if (bucket.count > majority.count) majority = bucket
      }
      const total = present.length
      expected[slotKey] = formatNormalized(majority.value)
      for (const entry of normalized) {
        const value = entry.slots.get(slotKey)
        if (!value) {
          // Only cells that miss a majority-shaped slot are anomalies; a slot
          // carried by fewer than half the cells belongs to a minority shape.
          if (total >= Math.ceil(normalized.length / 2)) {
            anomalies.push({
              kind: 'structure-mismatch',
              cell: entry.id,
              message: `missing slot ${slotKey}; column expects ${expected[slotKey]}`,
              expected: expected[slotKey],
              actual: null,
              confidence: total > 0 ? majority.count / total : null,
            })
          }
        } else if (normalizedKey(value) !== normalizedKey(majority.value)) {
          anomalies.push({
            kind: 'reference-offset',
            cell: entry.id,
            message: `slot ${slotKey} deviates from column pattern: expected ${expected[slotKey]}, actual ${formatNormalized(value)}`,
            expected: expected[slotKey],
            actual: formatNormalized(value),
            confidence: total > 0 ? majority.count / total : null,
            slot: slotKey,
            expectedOffsets: { colOffset: majority.value.colOffset, rowOffset: majority.value.rowOffset },
            actualOffsets: { colOffset: value.colOffset, rowOffset: value.rowOffset },
          })
        }
      }
    }
    reports.push({ sheet, column, cellCount: entries.length, expected, anomalies })
  }
  return reports
}

/**
 * A numeric cell's content, as `readWorkbookCells` stringifies it.
 *
 * Grouped digits are accepted because a number stored as *text* — what a paste
 * from a web page or a CSV import leaves behind — keeps its separators, and such
 * a cell sitting inside a formula column is exactly the hardcode break this
 * check exists to find. A genuine number never arrives grouped: Excel keeps it
 * as a value, and the reader stringifies `1000`, not `1,000`.
 */
const NUMERIC_PATTERN = /^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:%|e[+-]?\d+)?$/i

export function detectHardcodeBreaks(cells: Record<string, string>): PatternAnomaly[] {
  const columns = new Map<string, Array<{ id: string; row: number; isFormula: boolean; content: string }>>()
  for (const [id, content] of Object.entries(cells)) {
    const trimmed = content.trim()
    if (!trimmed) continue
    let cell
    try {
      cell = parseCellId(id)
    } catch {
      continue
    }
    const isFormula = trimmed.startsWith('=')
    if (!isFormula && !NUMERIC_PATTERN.test(trimmed)) continue
    const key = `${cell.sheet}:${cell.column}`
    if (!columns.has(key)) columns.set(key, [])
    columns.get(key)!.push({ id, row: cell.row, isFormula, content: trimmed })
  }

  const anomalies: PatternAnomaly[] = []
  for (const entries of columns.values()) {
    const formulas = entries.filter((entry) => entry.isFormula)
    const values = entries.filter((entry) => !entry.isFormula)
    if (formulas.length < 2 || values.length === 0) continue
    // Subtotal-structured columns interleave summary rows with data rows;
    // numeric cells between SUBTOTAL rows are data, not hardcoded breaks.
    if (formulas.some((entry) => /SUBTOTAL\(/i.test(entry.content))) continue
    const minRow = Math.min(...formulas.map((entry) => entry.row))
    const maxRow = Math.max(...formulas.map((entry) => entry.row))
    for (const value of values) {
      if (value.row >= minRow && value.row <= maxRow) {
        anomalies.push({
          kind: 'hardcode-break',
          cell: value.id,
          message: `numeric value ${value.content} inside formula column (formula rows ${minRow}-${maxRow})`,
          expected: '=formula',
          actual: value.content,
          confidence: null,
        })
      }
    }
  }
  return anomalies
}

export function detectEmptyGaps(cells: Record<string, string>): PatternAnomaly[] {
  const formulaRows = new Map<string, number[]>()
  const knownCells = new Set<string>()
  for (const [id, content] of Object.entries(cells)) {
    const trimmed = content.trim()
    if (!trimmed) continue
    let cell
    try {
      cell = parseCellId(id)
    } catch {
      continue
    }
    knownCells.add(canonicalCellId(cell.sheet, cell.column, cell.row))
    if (trimmed.startsWith('=')) {
      const key = `${cell.sheet}:${cell.column}`
      if (!formulaRows.has(key)) formulaRows.set(key, [])
      formulaRows.get(key)!.push(cell.row)
    }
  }
  const anomalies: PatternAnomaly[] = []
  for (const [key, rows] of formulaRows) {
    const sorted = [...rows].sort((a, b) => a - b)
    const separator = key.indexOf(':')
    const sheet = key.slice(0, separator)
    const column = key.slice(separator + 1)
    for (let i = 1; i < sorted.length; i++) {
      if (sorted[i] - sorted[i - 1] === 2) {
        const gapRow = sorted[i - 1] + 1
        const gapId = canonicalCellId(sheet, column, gapRow)
        if (!knownCells.has(gapId)) {
          anomalies.push({
            kind: 'empty-gap',
            cell: gapId,
            message: `empty cell between formula rows ${sorted[i - 1]} and ${sorted[i]}`,
            expected: `=formula at ${gapId}`,
            actual: 'empty',
            confidence: 1,
          })
        }
      }
    }
  }
  return anomalies
}

/**
 * Derived from the shared error list rather than written by hand.
 *
 * The hand-written version was `/#(?:REF|DIV\/0|VALUE|NAME\?|N\/A|NULL|NUM)!/g`,
 * which demanded a trailing `!` and therefore missed `#N/A` and `#NAME?` — the
 * errors a failed lookup and a misspelled function produce, and the two most
 * likely to show up in a workbook this check exists to police.
 */
const ERROR_TOKEN = new RegExp(`^(?:${excelErrorPattern()})$`)

/** A formula shown together with its error result, e.g. `=VLOOKUP(…) #N/A`. */
const TRAILING_ERROR_TOKEN = new RegExp(`\\s((?:${excelErrorPattern()}))$`)

/** `readWorkbookCells` serialises an error cell as `{"error":"#REF!"}`. */
const SERIALISED_ERROR = /^\{\s*"error"\s*:\s*"([^"]+)"\s*\}$/

/** String literals inside a formula, so an error *mentioned* in one is not a value. */
const FORMULA_STRING_LITERAL = /"(?:[^"\\]|\\.)*"/g

/**
 * Return the error token a cell holds, or null when the cell holds no error.
 *
 * Matching the token *anywhere* in the content — which is what an unanchored
 * regex did — flagged ordinary content as a broken cell: a note reading
 * `备注：#REF! 已修复`, and worse, the perfectly valid
 * `=IFERROR(A1/B1,"#N/A")`, whose entire purpose is to handle that error. It
 * also made this check disagree with `excel_find_errors`, which reads the typed
 * `cell.value.error` and therefore never over-reports.
 *
 * Only a few shapes actually occur: the value *is* the token (optionally as the
 * body of a formula, `=#REF!`); the token wrapped in the JSON envelope
 * `readWorkbookCells` produces; or a formula carrying its error result after the
 * formula text. The last is allowed only after string literals are removed, so a
 * formula that merely mentions an error inside a literal stays clean.
 */
function errorTokenOf(content: string): string | null {
  const trimmed = content.trim()
  if (!trimmed) return null
  const isFormula = trimmed.startsWith('=')
  const body = isFormula ? trimmed.slice(1).trim() : trimmed
  if (ERROR_TOKEN.test(body)) return body
  const serialised = SERIALISED_ERROR.exec(trimmed)
  if (serialised) return serialised[1]!
  if (isFormula) {
    const trailing = TRAILING_ERROR_TOKEN.exec(body.replace(FORMULA_STRING_LITERAL, ' '))
    if (trailing) return trailing[1]!
  }
  return null
}

/**
 * Detect cells whose content carries an Excel error value such as #REF! or
 * #DIV/0! (both literal error constants and formulas whose cached result is
 * an error token).
 */
export function detectErrorValues(cells: Record<string, string>): PatternAnomaly[] {
  const anomalies: PatternAnomaly[] = []
  for (const [id, content] of Object.entries(cells)) {
    const trimmed = content.trim()
    if (!trimmed) continue
    const token = errorTokenOf(trimmed)
    if (token !== null) {
      anomalies.push({
        kind: 'error-value',
        cell: id,
        message: `cell contains Excel error ${token}`,
        expected: 'valid value',
        actual: trimmed.slice(0, 200),
        confidence: 1,
      })
    }
  }
  return anomalies
}
