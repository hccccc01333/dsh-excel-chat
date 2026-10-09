export * from './operation-types.ts';
import type { ExcelOperation, ApplyOperationsResult, OperateResult } from './operation-types.ts';
export { findSheet, qualifySheetName, shiftFormulaReferences } from './op/core.ts';
export declare function applyOperationsToWorkbook(inputPath: string, operations: ExcelOperation[], outputPath: string): Promise<ApplyOperationsResult>;
export declare function operateWorkbookFile(path: string, operations: ExcelOperation[], outputPath: string): Promise<OperateResult>;
//# sourceMappingURL=operations.d.ts.map