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
 * Anything that needs to recognise an error value reads it from here.
 */
export const EXCEL_ERROR_VALUES = [
    '#DIV/0!',
    '#N/A',
    '#NAME?',
    '#NULL!',
    '#NUM!',
    '#REF!',
    '#VALUE!',
    '#GETTING_DATA',
];
/**
 * A regex source matching any error value, for callers that only have a string.
 *
 * Values are escaped (they contain `?` and `/`) and ordered longest first, so no
 * value can be shadowed by a shorter one that happens to be its prefix.
 */
export function excelErrorPattern() {
    return [...EXCEL_ERROR_VALUES]
        .sort((left, right) => right.length - left.length)
        .map((value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
        .join('|');
}
