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
 */
export declare function buildWorkbookInsight(path: string, sheet?: string): Promise<WorkbookInsight>;
//# sourceMappingURL=insight.d.ts.map