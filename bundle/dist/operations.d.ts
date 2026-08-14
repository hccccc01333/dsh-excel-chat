import ExcelJS from 'exceljs';
import { type ValidationResult } from './validator.ts';
export type ExcelOperation = {
    op: 'set';
    cells: Record<string, string>;
} | {
    op: 'fill';
    source: string;
    target: string;
} | {
    op: 'insertRows';
    sheet: string;
    row: number;
    count: number;
} | {
    op: 'deleteRows';
    sheet: string;
    row: number;
    count: number;
} | {
    op: 'addSheet';
    name: string;
} | {
    op: 'renameSheet';
    oldName: string;
    newName: string;
} | {
    op: 'deleteSheet';
    name: string;
} | {
    op: 'clear';
    cells: string[];
} | {
    op: 'merge';
    range: string;
} | {
    op: 'unmerge';
    range: string;
};
export interface OperationWarning {
    op: number;
    message: string;
}
export interface ApplyOperationsResult {
    warnings: OperationWarning[];
}
export interface OperateResult extends ApplyOperationsResult {
    outputPath: string;
    validation: ValidationResult;
}
export declare function findSheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined;
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
}): string;
export declare function applyOperationsToWorkbook(inputPath: string, operations: ExcelOperation[], outputPath: string): Promise<ApplyOperationsResult>;
export declare function operateWorkbookFile(path: string, operations: ExcelOperation[], outputPath: string): Promise<OperateResult>;
//# sourceMappingURL=operations.d.ts.map