export { EXCEL_ERROR_VALUES } from './error-values.ts';
export interface ErrorCell {
    /** Sheet-qualified cell id, e.g. "订单!D4". */
    cell: string;
    /** The error value, e.g. "#DIV/0!". */
    error: string;
    /** The formula that produced it, when the cell holds one. */
    formula: string | null;
}
export interface ErrorScan {
    total: number;
    /** Per-error-code counts, e.g. { "#DIV/0!": 2, "#N/A": 1 }. */
    counts: Record<string, number>;
    errorCells: ErrorCell[];
    sheetsScanned: string[];
}
/**
 * Find cells whose value is an Excel *error* rather than text that merely looks
 * like one.
 *
 * This reads the workbook directly instead of going through readWorkbookCells:
 * that helper stringifies every value, which would make a literal "#N/A" typed
 * by a user indistinguishable from a genuine error — and flagging those would
 * turn the check into noise.
 */
export declare function findErrorCells(data: Uint8Array, sheetFilter?: string): Promise<ErrorScan>;
//# sourceMappingURL=audit.d.ts.map