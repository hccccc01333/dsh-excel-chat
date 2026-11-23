/**
 * Build a native pivot table from a source range without Microsoft Excel.
 *
 * ExcelJS has no pivot model and drops pivot parts when it rewrites a workbook,
 * so the actual OOXML is emitted by xml-postprocess. This module computes the
 * aggregated cells (so the file is readable before Excel refreshes anything)
 * and describes the cache/table parts that need to be injected.
 */
import type ExcelJS from 'exceljs';
import type { PivotSpec } from '../xml-postprocess.ts';
import type { ExcelOperation } from '../operation-types.ts';
type PivotOperation = Extract<ExcelOperation, {
    op: 'pivot';
}>;
export declare function buildPivot(workbook: ExcelJS.Workbook, operation: PivotOperation, cacheId: number, partSuffix: number): PivotSpec;
export {};
//# sourceMappingURL=pivot-table.d.ts.map