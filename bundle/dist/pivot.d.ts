export interface PivotValueSpec {
    column: string;
    function: 'sum' | 'count' | 'average' | 'max' | 'min';
}
export interface PivotOptions {
    sheet: string;
    /** Source data range including the header row, e.g. "订单!A1:F7". */
    range: string;
    rows: string[];
    values: PivotValueSpec[];
    outputSheet?: string;
}
export interface PivotResult {
    pivotSheet: string;
    groups: number;
    recordCount: number;
}
/**
 * Create a native Excel pivot table by driving Excel COM (Windows): the cache
 * and pivot table are produced by Excel itself, so the output is always valid.
 */
export declare function createPivotTable(inputPath: string, options: PivotOptions, outPath: string): Promise<PivotResult>;
//# sourceMappingURL=pivot.d.ts.map