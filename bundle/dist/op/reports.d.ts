/**
 * Report-shaped operations: subtotals, the one-shot report, role presets and the
 * live-formula aggregate table.
 *
 * These share a shape — group by a column, aggregate others, write a summary sheet
 * whose cells are SUMIFS rather than values — so they are read together or not at
 * all. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation } from '../operation-types.ts';
export declare function applySubtotal(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'subtotal';
}>): number;
/**
 * One-shot report template: sort, subtotals, a dynamic SUMIFS summary sheet,
 * auto filter, header style, frozen header, and optional number format.
 * Ordering matters: subtotals run before the summary so its SUMIFS ranges
 * already cover the final data block (subtotal rows do not match group keys).
 */
export declare function applyReport(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'report';
}>): void;
/**
 * Role-based one-shot preset: 运营 gets a report with data bars, 产品 and 数分
 * get a report with color scales, and 数分 additionally writes a filtered copy.
 */
export declare function applyPreset(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'preset';
}>): void;
export declare function applyAggregateReport(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'aggregateReport';
}>): void;
//# sourceMappingURL=reports.d.ts.map