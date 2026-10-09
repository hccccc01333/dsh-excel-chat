/**
 * Inserting and deleting rows and columns, and keeping every formula honest.
 *
 * Excel shifts references when rows are inserted or deleted; nothing else in the
 * workbook does. These helpers walk the formulas and move the references, and mark
 * the ones that pointed into a deleted range so they read as errors instead of
 * silently pointing somewhere new.
 */
import ExcelJS from 'exceljs';
import type { OperationWarning } from '../operation-types.ts';
/** Delete rows with the same reference-shift semantics as the deleteRows op. */
export declare function deleteRowsFromSheet(workbook: ExcelJS.Workbook, sheetName: string, start: number, count: number, warnings: OperationWarning[], opIndex: number): void;
/** Delete columns with the same reference-shift semantics as the deleteColumns op. */
export declare function deleteColumnsFromSheet(workbook: ExcelJS.Workbook, sheetName: string, column: number, count: number, warnings: OperationWarning[], opIndex: number): void;
export declare function shiftWorkbookRows(workbook: ExcelJS.Workbook, editedSheet: string, threshold: number, rowDelta: number): void;
export declare function collectDeletedRangeRefs(workbook: ExcelJS.Workbook, editedSheet: string, start: number, end: number): string[];
export declare function shiftWorkbookColumns(workbook: ExcelJS.Workbook, editedSheet: string, threshold: number, colDelta: number): void;
export declare function collectDeletedColumnRefs(workbook: ExcelJS.Workbook, editedSheet: string, start: number, end: number): string[];
export declare function markDeletedRowRefs(workbook: ExcelJS.Workbook, editedSheet: string, start: number, end: number): void;
export declare function markDeletedColumnRefs(workbook: ExcelJS.Workbook, editedSheet: string, start: number, end: number): void;
//# sourceMappingURL=structure.d.ts.map