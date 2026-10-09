/**
 * Appearance: styles, data validation, conditional formatting and real tables.
 *
 * ExcelJS and the plugin describe styles differently (hAlign vs horizontal, fill as
 * a colour vs a pattern), so the conversion lives here in one place rather than at
 * each call site. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation, ExcelStyle } from '../operation-types.ts';
export declare function applyDataValidation(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'dataValidation';
}>): void;
export declare function looksLikeRange(value: string): boolean;
export declare function applyConditionalFormatting(workbook: ExcelJS.Workbook, range: string, rules: Extract<ExcelOperation, {
    op: 'conditionalFormatting';
}>['rules']): void;
export declare function excelStyleToWorkbookStyle(style: ExcelStyle): ExcelJS.Style;
export declare function addTable(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'addTable';
}>): void;
export declare function applyStyle(workbook: ExcelJS.Workbook, range: string, style: ExcelStyle): void;
export declare function normalizeColor(color: string): string;
//# sourceMappingURL=formatting.d.ts.map