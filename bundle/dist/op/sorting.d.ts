/**
 * Sorting a range, including the keys that are not values.
 *
 * A sort key can be a value, a fill colour or a font colour, and a custom list
 * overrides the natural order — so the comparison needs the cell, not just its
 * text. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation } from '../operation-types.ts';
export declare function sortRange(workbook: ExcelJS.Workbook, range: string, keys: Extract<ExcelOperation, {
    op: 'sortRange';
}>['keys'], headerRows: number): void;
//# sourceMappingURL=sorting.d.ts.map