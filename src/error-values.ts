/**
 * The error values Excel can store in a cell — the single source of truth.
 *
 * This list used to exist twice: here (for `excel_find_errors`) and again as a
 * hand-written regex inside `patterns.ts` (for the post-edit health check).
 * They drifted, and the copy in `patterns.ts` had quietly lost `#N/A` and
 * `#NAME?` — its regex demanded a trailing `!`, which neither value has. The
 * result was that a workbook full of failed lookups passed the health check
 * while `excel_find_errors` reported the very same cells.
 *
 * It was also short. `findErrorCells` never consulted this list — it accepts
 * whatever error token the file carries — so it reported `#SPILL!` and friends
 * while the health check, which does derive from the list, could not. Every
 * token Excel is documented to write is therefore listed here.
 *
 * Anything that needs to recognise an error value reads it from here.
 */
export const EXCEL_ERROR_VALUES = [
  // The seven Excel has always had.
  '#DIV/0!',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#REF!',
  '#VALUE!',
  // Legacy external-data error.
  '#GETTING_DATA',
  // Excel 365 dynamic arrays and linked data types.
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#UNKNOWN!',
  '#CONNECT!',
  '#BUSY!',
  // Python in Excel.
  '#PYTHON!',
  '#TIMEOUT!',
] as const

/**
 * A regex source matching any error value, for callers that only have a string.
 *
 * Values are escaped (they contain `?` and `/`) and ordered longest first, so no
 * value can be shadowed by a shorter one that happens to be its prefix.
 */
export function excelErrorPattern(): string {
  return [...EXCEL_ERROR_VALUES]
    .sort((left, right) => right.length - left.length)
    .map((value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
    .join('|')
}
