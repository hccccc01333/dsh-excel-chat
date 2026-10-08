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
/**
 * The inverse of `cellContent`: turn a serialized cell string back into a typed
 * value. Shared so that every writer — `set`, `importCsv`, and the patch-log
 * rollback behind `excel_undo` — infers the same types from the same strings.
 *
 * The rollback used to assign the raw string, which turned every restored
 * number, boolean and date into *text*: undoing an edit left `42` as the string
 * "42". The damage was invisible because `cellContent` renders both the same
 * way, so comparing the strings showed no difference at all.
 */
export declare function contentToCellValue(content: string): ExcelJS.CellValue;
/**
 * `cellContent` wraps an error cell as `{"error":"#REF!"}` so an error stays
 * distinguishable from text that happens to spell the same token. A consumer
 * that wants the plain representation — a CSV, which Excel writes as the bare
 * token — unwraps it here instead of re-deriving the shape.
 */
export declare function plainContent(content: string): string;
/**
 * Sheet names in workbook order. The plan salvage needs them to fill in a
 * missing `sheet`/range prefix, so anything applying model-supplied operations
 * has to know the sheets before it can validate them.
 */
export declare function readWorkbookSheetNames(path: string): Promise<string[]>;
//# sourceMappingURL=workbook.d.ts.map