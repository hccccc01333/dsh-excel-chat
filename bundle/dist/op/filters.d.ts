/**
 * Selecting rows: filtering to a range, testing a criterion, and mail merge.
 *
 * The criterion matcher is shared by the filter and the merge so the two cannot
 * disagree about what `contains` means. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation } from '../operation-types.ts';
export declare function applyFilterToRange(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'filterToRange';
}>): void;
export declare function matchesCriterion(actual: ExcelJS.CellValue, operator: 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains', expected: string | number): boolean;
export declare function applyMailMerge(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'mailMerge';
}>): void;
//# sourceMappingURL=filters.d.ts.map