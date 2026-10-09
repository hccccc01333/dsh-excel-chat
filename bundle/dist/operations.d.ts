export * from './operation-types.ts';
import type { ExcelOperation, ApplyOperationsResult, OperateResult } from './operation-types.ts';
export { findSheet, qualifySheetName, shiftFormulaReferences } from './op/core.ts';
export declare function applyOperationsToWorkbook(inputPath: string, operations: ExcelOperation[], outputPath: string): Promise<ApplyOperationsResult>;
/** Visible text of a cell for width estimation: formula cells use their cached result. */
/**
 * Approximate display width in character units: CJK/fullwidth characters count
 * as 2 columns, everything else as 1.
 */
/** Join keys are compared trimmed + lowercased, numbers via their text form. */
/** Aggregations where a grand total of the computed grid is meaningful. */
/**
 * Fixed-width split: take `widths[i]` characters per output column. Anything
 * past the last width becomes a trailing column rather than being dropped.
 */
export declare function operateWorkbookFile(path: string, operations: ExcelOperation[], outputPath: string): Promise<OperateResult>;
//# sourceMappingURL=operations.d.ts.map