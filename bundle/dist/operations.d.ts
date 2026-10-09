import ExcelJS from 'exceljs';
export * from './operation-types.ts';
import type { ExcelOperation, ApplyOperationsResult, OperateResult } from './operation-types.ts';
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
    rowDeletedStart?: number;
    rowDeletedEnd?: number;
    colDeletedStart?: number;
    colDeletedEnd?: number;
}): string;
/**
 * Qualify a sheet name for embedding inside an Excel formula string. Names
 * that are not a plain identifier (spaces, punctuation, leading digits, CJK
 * beyond letters) must be single-quoted with internal quotes doubled, or the
 * generated reference breaks and Excel repairs/drops it.
 */
export declare function qualifySheetName(name: string): string;
export declare function applyOperationsToWorkbook(inputPath: string, operations: ExcelOperation[], outputPath: string): Promise<ApplyOperationsResult>;
export declare function operateWorkbookFile(path: string, operations: ExcelOperation[], outputPath: string): Promise<OperateResult>;
//# sourceMappingURL=operations.d.ts.map