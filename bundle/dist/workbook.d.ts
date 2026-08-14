import ExcelJS from 'exceljs';
import { type ValidationResult } from './validator.ts';
export declare function cellContent(cell: ExcelJS.Cell): string | null;
export declare function readWorkbookCells(data: Uint8Array): Promise<Record<string, string>>;
export declare function validateWorkbookFile(path: string): Promise<ValidationResult>;
//# sourceMappingURL=workbook.d.ts.map