/** Default page size suggested by `excel_profile` for chunked `excel_read` calls. */
export declare const PROFILE_PAGE_SIZE = 100;
export interface ProfileTopValue {
    value: string;
    count: number;
}
export interface ColumnProfile {
    column: string;
    header: string | null;
    dtype: 'string' | 'number' | 'date' | 'boolean' | 'mixed' | 'empty';
    nonEmpty: number;
    unique: number;
    uniqueCapped: boolean;
    missing: number;
    min?: number;
    max?: number;
    mean?: number;
    minDate?: string;
    maxDate?: string;
    topValues: ProfileTopValue[];
    samples: string[];
}
export interface SheetProfile {
    sheet: string;
    rowCount: number;
    columnCount: number;
    usedRange: string;
    headerRow: number | null;
    formulaCells: number;
    dataRows: number;
    profiledRows: number;
    truncated: boolean;
    readHint: string;
    columns: ColumnProfile[];
}
export interface WorkbookProfile {
    path: string;
    sheetCount: number;
    pageSize: number;
    sheets: SheetProfile[];
}
/**
 * Profile a workbook into a compact structural digest: sheet dimensions,
 * detected header row, per-column dtype/missing/unique/stats/top values, and
 * the range to read first. This is the "structured table encoding" that keeps
 * large sheets readable in one tool call instead of dumping every cell.
 */
export declare function profileWorkbook(path: string, sheet?: string): Promise<WorkbookProfile>;
//# sourceMappingURL=profile.d.ts.map