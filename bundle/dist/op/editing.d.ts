/**
 * Editing cells: filling, copying, merging, transposing, clearing, finding text.
 *
 * These are the operations that move or rewrite content in place, as opposed to the
 * ones that compute something. They share the reference-shifting question — does the
 * edit move other cells with it — which is why they read together. Split out of
 * `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { OperationWarning } from '../operation-types.ts';
export declare function applyFill(workbook: ExcelJS.Workbook, sourceId: string, targetRange: string): void;
export declare function copyRange(workbook: ExcelJS.Workbook, sourceRange: string, targetCell: string, move: boolean, valuesOnly?: boolean, warnings?: {
    push(w: OperationWarning): void;
}, opIndex?: number): void;
export declare function fillSeries(workbook: ExcelJS.Workbook, startId: string, targetRange: string, step?: number): void;
export declare function findReplace(workbook: ExcelJS.Workbook, find: string, replace: string, sheetName: string | undefined, matchCase: boolean): number;
export declare function duplicateSheet(workbook: ExcelJS.Workbook, name: string, newName: string): void;
export declare function renameSheetReferences(workbook: ExcelJS.Workbook, oldName: string, newName: string): void;
export declare function applyMerge(workbook: ExcelJS.Workbook, range: string, unmerge: boolean): void;
export declare function transposeRange(workbook: ExcelJS.Workbook, sourceRange: string, targetCell: string): void;
export declare function clearRange(workbook: ExcelJS.Workbook, range: string, mode: 'contents' | 'formats' | 'all'): void;
//# sourceMappingURL=editing.d.ts.map