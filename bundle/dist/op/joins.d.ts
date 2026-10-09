/**
 * Joining two sheets on a key column.
 *
 * The join is exact-match and writes values, not formulas — a VLOOKUP would break
 * the moment the lookup sheet is sorted, and the point of this operation is a
 * filled-in column that stays put. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation, OperationWarning } from '../operation-types.ts';
export declare function joinSheets(workbook: ExcelJS.Workbook, operation: Extract<ExcelOperation, {
    op: 'joinSheets';
}>, warnings: OperationWarning[], opIndex: number): void;
//# sourceMappingURL=joins.d.ts.map