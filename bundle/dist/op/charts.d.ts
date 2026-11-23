/**
 * Build the OOXML description of a chart from workbook ranges.
 *
 * ExcelJS 4.4 has no chart API at all — it even drops existing charts when a
 * workbook is rewritten — so the chart parts are emitted by xml-postprocess.
 * This module only reads cells and produces a sheet-independent description.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation } from '../operation-types.ts';
import type { ChartSpec } from '../xml-postprocess.ts';
type ChartOperation = Extract<ExcelOperation, {
    op: 'addChart';
}>;
export declare function buildChartSpec(workbook: ExcelJS.Workbook, operation: ChartOperation): ChartSpec;
export {};
//# sourceMappingURL=charts.d.ts.map