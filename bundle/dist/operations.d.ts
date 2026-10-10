export * from './operation-types.ts';
import type { ExcelOperation, ApplyOperationsResult, OperateResult } from './operation-types.ts';
export { findSheet, qualifySheetName, shiftFormulaReferences } from './op/core.ts';
/**
 * How to treat a workbook that contains parts ExcelJS cannot write back.
 *
 * `refuse` is the default because the alternative already cost someone their macros:
 * a warning is printed, the edit goes ahead, and the feature is gone from the file.
 * A warning that scrolls past in a tool result is not consent.
 *
 * There are two honest ways out, and both are a deliberate act:
 *   - write to a **different** output path, so the original is never touched
 *   - pass `allow`, which is the caller saying it knows what it is giving up
 */
export interface ApplyOperationsOptions {
    onUnpreserved?: 'refuse' | 'allow';
}
export declare function applyOperationsToWorkbook(inputPath: string, operations: ExcelOperation[], outputPath: string, options?: ApplyOperationsOptions): Promise<ApplyOperationsResult>;
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
export declare function operateWorkbookFile(path: string, operations: ExcelOperation[], outputPath: string, options?: ApplyOperationsOptions): Promise<OperateResult>;
//# sourceMappingURL=operations.d.ts.map