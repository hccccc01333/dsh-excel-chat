/**
 * Whole-sheet concerns: visibility, order, tab colour, hyperlinks, freeze panes.
 *
 * Everything here acts on the sheet as an object rather than on its cells, which is
 * why it reads as one group. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation } from '../operation-types.ts';
export declare function setHyperlink(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'setHyperlink';
}>): void;
export declare function copyStyle(workbook: ExcelJS.Workbook, sourceId: string, targetRange: string): void;
export declare function freezeFormulas(workbook: ExcelJS.Workbook, range: string): {
    frozen: number;
    skipped: number;
};
export declare function uniqueValues(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'uniqueValues';
}>): number;
export declare function applySheetView(sheet: ExcelJS.Worksheet, patch: (view: ExcelJS.WorksheetView) => void): void;
export declare function moveSheet(workbook: ExcelJS.Workbook, name: string, position: number): void;
export declare function applyRankColumn(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'rankColumn';
}>): void;
//# sourceMappingURL=sheets.d.ts.map