/**
 * Sorting a range, including the keys that are not values.
 *
 * A sort key can be a value, a fill colour or a font colour, and a custom list
 * overrides the natural order — so the comparison needs the cell, not just its
 * text. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs'
import { columnToNumber, numberToColumn } from '../formula.ts'
import type { ExcelOperation } from '../operation-types.ts'
import { parseRange } from './core.ts'
import { normalizeColor } from './formatting.ts'
export function sortRange(
  workbook: ExcelJS.Workbook,
  range: string,
  keys: Extract<ExcelOperation, { op: 'sortRange' }>['keys'],
  headerRows: number,
): void {
  const parsed = parseRange(workbook, range)
  for (const key of keys) {
    if ((key.by === 'fill' || key.by === 'font') && key.color === undefined) {
      throw new Error(`sortRange key with by: "${key.by}" requires color`)
    }
  }
  const keyColumns = keys.map((key) => ({
    column: columnToNumber(key.column),
    direction: key.direction ?? 'asc',
    by: key.by ?? 'value',
    color: key.color === undefined ? undefined : normalizeColor(key.color),
    customList: key.customList,
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
    styles: Record<string, Partial<ExcelJS.Style>>
    keys: Array<string | number | Date | null>
  }
  const rows: SortRow[] = []
  for (let row = parsed.startRow + headerRows; row <= parsed.endRow; row++) {
    const cells: Record<string, ExcelJS.CellValue> = {}
    const styles: Record<string, Partial<ExcelJS.Style>> = {}
    const keyValues: SortRow['keys'] = []
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const letter = numberToColumn(col)
      const cell = parsed.sheet.getCell(`${letter}${row}`)
      cells[letter] = cell.value
      // Carry formatting with the row, the way Excel does — without this a
      // sort by fill colour would move the values out from under the colours.
      styles[letter] = cell.style
      const key = keyColumns.find((candidate) => candidate.column === col)
      if (key) keyValues.push(sortKeyOf(cell, key))
    }
    rows.push({ cells, styles, keys: keyValues })
  }
  rows.sort((a, b) => compareSortRows(a.keys, b.keys, keyColumns.map((key) => key.direction)))
  for (let index = 0; index < rows.length; index++) {
    const targetRow = parsed.startRow + headerRows + index
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      const letter = numberToColumn(col)
      const target = parsed.sheet.getCell(`${letter}${targetRow}`)
      target.value = rows[index]!.cells[letter] ?? null
      target.style = (rows[index]!.styles[letter] ?? {}) as ExcelJS.Style
    }
  }
}

/** Sort keys are pre-reduced to plain scalars so the comparator stays dumb. */
type SortKeyColumn = {
  column: number
  direction: 'asc' | 'desc'
  by: 'value' | 'fill' | 'font'
  color?: string
  customList?: string[]
}

/**
 * Reduce a cell to whatever its sort key should compare on.
 *
 * Colour keys collapse to a 0/1 group — cells carrying the requested colour sort
 * first, everything else follows — which is what Excel's "move the selected
 * colour to the top" does. Custom lists become the value's position in the list,
 * with unlisted values pushed past the end so they trail the listed ones. Both
 * therefore ride the existing value comparator with no special cases in it.
 */
function sortKeyOf(cell: ExcelJS.Cell, key: SortKeyColumn): string | number | Date | null {
  if (key.by === 'fill' || key.by === 'font') {
    const cellColor = key.by === 'fill' ? fillColorOf(cell) : fontColorOf(cell)
    return cellColor !== null && cellColor === key.color ? 0 : 1
  }
  if (key.customList !== undefined) {
    const index = key.customList.indexOf(String(cell.value ?? ''))
    return index >= 0 ? index : key.customList.length
  }
  return cell.value as string | number | Date | null
}

function fillColorOf(cell: ExcelJS.Cell): string | null {
  const fill = cell.style?.fill as { type?: string; fgColor?: { argb?: string } } | undefined
  if (fill?.type !== 'pattern') return null
  return typeof fill.fgColor?.argb === 'string' ? fill.fgColor.argb.toUpperCase() : null
}

function fontColorOf(cell: ExcelJS.Cell): string | null {
  const argb = cell.font?.color?.argb
  return typeof argb === 'string' ? argb.toUpperCase() : null
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

