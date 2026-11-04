/**
 * Layout math for the Excel toolview, kept out of the React component.
 *
 * These are the parts that fail silently rather than loudly — A1 addresses,
 * render caps, the payload x-data-spreadsheet expects — and unlike the view
 * itself they need neither a DOM nor a renderer to exercise.
 */

export interface TableCellShape {
  id: string
  value: unknown
  formula?: string
  type?: string
}

export interface TableSheetShape {
  sheet: string
  cells: TableCellShape[]
}

/** Shape of an excel_read / excel_preview result. */
export interface TableShape {
  sheets?: TableSheetShape[]
  markdown?: string
}

export interface GridSheetData {
  name: string
  rows: Record<number, { cells: Record<number, { text: string }> }>
}

/**
 * Display caps: bigger sheets get a truncation notice instead of silent loss.
 *
 * Fullscreen raises them. The notice quotes the numbers from `measureSheets`
 * rather than constants of its own, so what it promises and what the grid does
 * cannot drift apart.
 */
export const INLINE_RENDER_ROWS = 2000
export const INLINE_RENDER_COLS = 100
export const FULLSCREEN_RENDER_ROWS = 20000
export const FULLSCREEN_RENDER_COLS = 500

/** Minimum grid size, so a one-cell sheet still looks like a worksheet. */
const MIN_RENDER_ROWS = 10
const MIN_RENDER_COLS = 5

/** 1-based column index → letters ("A", "Z", "AA"), matching Excel addressing. */
export function columnName(index: number): string {
  let name = ''
  let value = index
  while (value > 0) {
    const remainder = (value - 1) % 26
    name = `${String.fromCharCode(65 + remainder)}${name}`
    value = Math.floor((value - remainder - 1) / 26)
  }
  return name
}

/** A cell's displayed text: the formula when there is one, else the literal. */
export function cellText(value: unknown, formula?: string): string {
  if (formula) return formula
  if (value === null || value === undefined) return ''
  if (typeof value === 'object') return JSON.stringify(value)
  return String(value)
}

/** "订单!AB12" → { row: 11, col: 27 } (0-based); null when not an A1 address. */
export function parseCellId(id: string): { row: number; col: number } | null {
  const body = id.split('!').pop() ?? ''
  const match = /^([A-Za-z]+)(\d+)$/.exec(body)
  if (match === null) return null
  const row = Number(match[2])
  if (!Number.isFinite(row) || row < 1) return null
  const col = [...match[1]!.toUpperCase()]
    .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0)
  return { row, col }
}

export interface SheetLayout {
  /** Grid size handed to the library (already clamped, floored at 10 × 5). */
  rows: number
  cols: number
  /** The cap this mode allows — the truncation notice quotes these. */
  capRows: number
  capCols: number
  /** True when the sheet itself is bigger than the cap. */
  truncated: boolean
}

/** Grid size for `sheets`: content size clamped to the cap for this mode. */
export function measureSheets(sheets: TableSheetShape[], fullscreen: boolean): SheetLayout {
  const capRows = fullscreen ? FULLSCREEN_RENDER_ROWS : INLINE_RENDER_ROWS
  const capCols = fullscreen ? FULLSCREEN_RENDER_COLS : INLINE_RENDER_COLS
  let maxRows = 1
  let maxCols = 1
  for (const sheet of sheets) {
    for (const cell of sheet.cells) {
      const address = parseCellId(cell.id)
      if (address === null) continue
      if (address.row > maxRows) maxRows = address.row
      if (address.col > maxCols) maxCols = address.col
    }
  }
  return {
    rows: Math.min(capRows, Math.max(MIN_RENDER_ROWS, maxRows)),
    cols: Math.min(capCols, Math.max(MIN_RENDER_COLS, maxCols)),
    capRows,
    capCols,
    truncated: maxRows > capRows || maxCols > capCols,
  }
}

/** x-data-spreadsheet payload: row/col keyed cells with display text, per sheet. */
export function toGridData(sheets: TableSheetShape[]): GridSheetData[] {
  return sheets.map((sheet) => {
    const rows: GridSheetData['rows'] = {}
    for (const cell of sheet.cells) {
      const address = parseCellId(cell.id)
      if (address === null) continue
      const row = (rows[address.row - 1] ??= { cells: {} })
      row.cells[address.col - 1] = { text: cellText(cell.value, cell.formula) }
    }
    return { name: sheet.sheet, rows }
  })
}
