/**
 * The two-dimensional cross-tab: row dimension × column dimension, cells live SUMIFS.
 *
 * Unlike a pivot table, every cell here is a formula, so the result updates when the
 * source does. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation, OperationWarning } from '../operation-types.ts';
export declare function applyCrosstab(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'crosstab';
}>, warnings: OperationWarning[], opIndex: number): void;
/** Formats exceljs can embed, and the size ceiling that keeps memory sane. */
/**
 * Embed an image anchored at a cell. Cross-platform: exceljs writes the media
 * part and the drawing XML itself, so no Excel installation is involved. Because
 * we mutate the loaded workbook (rather than rebuilding it), images that were
 * already in the file survive — verified by reading the media parts back.
 */
/** Resolve image bytes plus exceljs's extension token, from a path or base64. */
/**
 * The rendered size for an embedded image.
 *
 * Both sides given → use them as-is. Neither → the image's own pixel size. Just
 * one → scale the other from the intrinsic size, so asking for `width: 300` on a
 * 100×200 image renders 300×600 rather than squashing it to 300×100. Returns
 * null only when the caller gave a single side and the header is unreadable.
 */
/** Proportional counterpart to a given dimension, never rounding down to zero. */
//# sourceMappingURL=crosstab.d.ts.map