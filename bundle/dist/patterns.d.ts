export type PatternAnomalyKind = 'reference-offset' | 'structure-mismatch' | 'hardcode-break' | 'empty-gap' | 'circular-reference';
export interface PatternAnomaly {
    kind: PatternAnomalyKind;
    cell: string;
    message: string;
    expected: string | null;
    actual: string | null;
    confidence: number | null;
    slot?: string;
    expectedOffsets?: {
        colOffset: number | null;
        rowOffset: number | null;
    };
    actualOffsets?: {
        colOffset: number | null;
        rowOffset: number | null;
    };
}
export interface ColumnPatternReport {
    sheet: string;
    column: string;
    cellCount: number;
    expected: Record<string, string>;
    anomalies: PatternAnomaly[];
}
export declare function detectPatternAnomalies(cells: Record<string, string>): ColumnPatternReport[];
export declare function detectHardcodeBreaks(cells: Record<string, string>): PatternAnomaly[];
export declare function detectEmptyGaps(cells: Record<string, string>): PatternAnomaly[];
//# sourceMappingURL=patterns.d.ts.map