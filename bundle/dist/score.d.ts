import type { CellDiffEntry } from './diff.ts';
export interface WorkbookScore {
    total: number;
    matched: number;
    mismatched: number;
    accuracy: number;
    passes: boolean;
    mismatches: CellDiffEntry[];
}
export declare function normalizeCellId(id: string): string;
/** Formula comparison tolerance: case, whitespace, and numeric formatting. */
export declare function cellValueEquals(a: string | null, b: string | null): boolean;
/**
 * Compare a candidate workbook against an oracle workbook cell by cell.
 * Passes when every cell matches (formula case/whitespace and numeric
 * formatting are tolerated). The mismatch entries use the same shape as
 * diffCellMaps(oracle, candidate), so added cells are candidate extras.
 */
export declare function scoreWorkbookAgainstOracle(candidate: Record<string, string>, oracle: Record<string, string>): WorkbookScore;
export declare function scoreWorkbookFiles(candidatePath: string, oraclePath: string): Promise<WorkbookScore>;
//# sourceMappingURL=score.d.ts.map