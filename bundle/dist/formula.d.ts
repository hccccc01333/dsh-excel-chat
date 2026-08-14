export interface RefPoint {
    sheet: string | null;
    column: string;
    row: number | null;
    absColumn: boolean;
    absRow: boolean;
}
export interface ParsedRef {
    start: RefPoint;
    end: RefPoint | null;
    /** Text range of the reference inside the formula (without the leading "="). */
    range: {
        start: number;
        end: number;
    };
}
export interface ParsedFormula {
    raw: string;
    references: ParsedRef[];
}
export declare const DEFAULT_SHEET = "Sheet1";
export declare function normalizeSheet(sheet: string | null): string;
export declare function canonicalCellId(sheet: string | null, column: string, row: number): string;
export declare function columnToNumber(column: string): number;
export declare function numberToColumn(value: number): string;
export declare function parseFormula(input: string): ParsedFormula;
export interface ParsedCellId {
    sheet: string;
    column: string;
    row: number;
}
export declare function parseCellId(id: string): ParsedCellId;
/**
 * Shift every relative row reference in a formula by rowDelta, preserving
 * columns, absolute rows ($4), whole-column references, and sheet prefixes.
 * Returns the original formula when any shift would leave the sheet (row < 1).
 */
export declare function shiftFormulaRow(formula: string, rowDelta: number): string;
//# sourceMappingURL=formula.d.ts.map