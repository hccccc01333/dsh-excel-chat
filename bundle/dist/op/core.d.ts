/**
 * Cells, ranges and the small helpers every operation needs.
 *
 * The bottom of the stack: `operations.ts` used to hold all of this alongside the
 * dispatcher and 58 handlers. Nothing here calls anything above it, which is what
 * lets the rest be split without cycles.
 */
import ExcelJS from 'exceljs';
export declare const RANGE_LINE: RegExp;
export declare function findSheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined;
export declare function resolveCell(workbook: ExcelJS.Workbook, id: string): ExcelJS.Cell;
export declare function writeContent(cell: ExcelJS.Cell, content: unknown): void;
/**
 * Convert user-provided text into an Excel value: formulas stay formulas,
 * plain numbers/dates/booleans keep their type, everything else is text.
 * Workplace spreadsheets break when "100" is written as text, so numeric
 * strings are typed before they reach ExcelJS. The conversion itself lives next
 * to `cellContent` in workbook.ts so every writer shares it.
 */
interface ParsedRange {
    sheet: ExcelJS.Worksheet;
    startCol: number;
    startRow: number;
    endCol: number;
    endRow: number;
}
export declare function parseRange(workbook: ExcelJS.Workbook, range: string): ParsedRange;
/**
 * Shift selected reference points of a formula. rowDelta/colDelta apply to
 * relative rows/columns; rowThreshold/colThreshold gate the shift so row edits
 * only move references at or below the insertion/deletion point. When
 * editedSheet is set, only references pointing into that sheet are shifted.
 */
export declare function shiftFormulaReferences(formula: string, baseSheet: string, editedSheet: string | null, options?: {
    rowDelta?: number;
    colDelta?: number;
    rowThreshold?: number;
    colThreshold?: number;
    rowDeletedStart?: number;
    rowDeletedEnd?: number;
    colDeletedStart?: number;
    colDeletedEnd?: number;
}): string;
export declare function cellContentOf(cell: ExcelJS.Cell): string;
/**
 * Qualify a sheet name for embedding inside an Excel formula string. Names
 * that are not a plain identifier (spaces, punctuation, leading digits, CJK
 * beyond letters) must be single-quoted with internal quotes doubled, or the
 * generated reference breaks and Excel repairs/drops it.
 */
export declare function qualifySheetName(name: string): string;
/** Build an absolute A1 column range for embedding in a formula: Sheet!$C$2:$C$9. */
export declare function absoluteColumnRef(sheetName: string, colLetter: string, fromRow: number, toRow: number): string;
/** Delete rows with the same reference-shift semantics as the deleteRows op. */
export declare function properCase(text: string): string;
/** Fullwidth ASCII/space/punctuation to halfwidth, then trim and collapse spaces. */
export declare function normalizeTextValue(text: string): string;
/**
 * Normalized similarity in [0, 1] for fuzzy matching: exact match is 1,
 * otherwise 1 minus the Levenshtein distance ratio over the longer string.
 * Callers normalize (trim/lowercase) before calling.
 */
export declare function similarity(a: string, b: string): number;
export declare function parseTargetCell(workbook: ExcelJS.Workbook, target: string, defaultSheet: string): {
    sheet: ExcelJS.Worksheet;
    col: number;
    row: number;
};
export declare function splitByWidth(text: string, widths: number[]): string[];
export {};
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
//# sourceMappingURL=core.d.ts.map