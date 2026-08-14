import type { ColumnTable } from './ir.ts';
/**
 * Detect a table schema from cell content: the first sheet row with at least
 * two non-numeric, non-formula cells becomes the header row. Text cells map
 * to their column letters. Returns null when no header row is found.
 */
export declare function detectTableFromCells(cells: Record<string, string>, sheetName?: string): ColumnTable | null;
//# sourceMappingURL=tables.d.ts.map