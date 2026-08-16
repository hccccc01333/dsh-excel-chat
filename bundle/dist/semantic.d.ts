import { type ColumnProfile } from './profile.ts';
export type ColumnRole = 'time' | 'measure' | 'dimension' | 'id' | 'unknown';
export interface SemanticColumn {
    column: string;
    header: string | null;
    role: ColumnRole;
    dtype: ColumnProfile['dtype'];
}
export interface SemanticSheet {
    sheet: string;
    columns: SemanticColumn[];
    grain: string;
    derivedMetrics: string[];
}
export interface SemanticProfile {
    sheets: SemanticSheet[];
    joinKeys: Array<{
        left: string;
        right: string;
        key: string;
    }>;
    summary: string;
}
/**
 * Workbook semantic layer (v0.36 first slice): classify each column's role
 * (time / measure / dimension / id), derive sheet grain and formula-based
 * metrics, and find cross-sheet join keys. The summary feeds the planner and
 * verifier so the agent stops guessing "is 地区 column B?".
 */
export declare function buildWorkbookSemanticProfile(path: string, sheet?: string): Promise<SemanticProfile>;
//# sourceMappingURL=semantic.d.ts.map