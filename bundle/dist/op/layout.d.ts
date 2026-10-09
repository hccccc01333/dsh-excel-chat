/**
 * Page setup and the text metrics that drive column widths.
 *
 * `displayWidth` counts CJK characters as two columns, which is what makes an
 * autofit look right in a sheet full of Chinese; it lives with page setup because
 * both are about how the sheet presents rather than what it contains. Split out of
 * `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation } from '../operation-types.ts';
export declare function applyPageSetup(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'pageSetup';
}>): void;
export declare function displayTextOf(cell: ExcelJS.Cell): string;
export declare function displayWidth(text: string): number;
//# sourceMappingURL=layout.d.ts.map