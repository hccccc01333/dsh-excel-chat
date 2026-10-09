/**
 * Embedding images, including reading their intrinsic size from the bytes.
 *
 * The image parts are written at the XML layer rather than through ExcelJS, so the
 * extent maths lives here: a caller may give a pixel width, a height, both or
 * neither, and the aspect ratio has to survive all four cases. Split out of
 * `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation } from '../operation-types.ts';
export declare function insertImageIntoSheet(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'insertImage';
}>): Promise<void>;
//# sourceMappingURL=images.d.ts.map