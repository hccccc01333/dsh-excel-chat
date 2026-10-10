export type InsightSeverity = 'info' | 'warn' | 'alert';
export interface InsightFinding {
    severity: InsightSeverity;
    category: string;
    message: string;
}
export interface SheetInsight {
    sheet: string;
    summary: string;
    findings: InsightFinding[];
}
export interface WorkbookInsight {
    summary: string;
    sheets: SheetInsight[];
    suggestions: string[];
}
/**
 * Heuristic data insight report (ExcelGenius2-style "upload -> summary +
 * anomalies"): per-sheet one-liner, missing/duplicate/outlier/normalization
 * findings, and concrete next-step suggestions. Deterministic, no LLM needed.
 *
 * Two kinds of finding come back. Most answer "is this data trustworthy" — missing
 * values, duplicates, outliers. The rest come from `analysis.ts` and answer "what do the
 * numbers say" — a trend across periods, a measure concentrated in a few categories.
 * The second kind was missing entirely: a sheet with a clean 100→260 ramp and one region
 * four times the size of the others produced no findings at all.
 */
export declare function buildWorkbookInsight(path: string, sheet?: string): Promise<WorkbookInsight>;
//# sourceMappingURL=insight.d.ts.map