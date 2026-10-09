/**
 * CSV import and export, kept beside the workbook operations they are invoked from.
 *
 * Both go through `csv.ts`, which owns the quoting rules and the formula-injection
 * guard; this module is only the part that touches a workbook. Split out of
 * `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation } from '../operation-types.ts';
export declare function importCsv(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'importCsv';
}>): Promise<void>;
export declare function exportCsv(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'exportCsv';
}>): Promise<void>;
//# sourceMappingURL=csv.d.ts.map