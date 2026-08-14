import ExcelJS from 'exceljs';
import { type ValidationResult } from './validator.ts';
export declare function cellContent(cell: ExcelJS.Cell): string | null;
/**
 * ExcelJS crashes when a worksheet's `<tableParts>` points at a pivot table
 * (it only understands regular tables). Strip those anchors before loading so
 * read/validate/operate keep working on files that contain pivot tables.
 * Pivot parts are dropped on rewrite — acceptable, since ExcelJS cannot
 * preserve them anyway.
 */
export declare function stripPivotTableParts(data: Uint8Array): Uint8Array;
export declare function readWorkbookCells(data: Uint8Array): Promise<Record<string, string>>;
export declare function validateWorkbookFile(path: string): Promise<ValidationResult>;
//# sourceMappingURL=workbook.d.ts.map