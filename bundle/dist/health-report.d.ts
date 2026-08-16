export declare const HEALTH_REPORT_SHEET = "_dsh_\u4F53\u68C0\u62A5\u544A";
export interface HealthReportResult {
    path: string;
    healthScore: number;
    formulaCount: number;
    anomalyCount: number;
    reportSheet: string;
    summary: string;
}
/**
 * Write a formula health report INTO the workbook itself: a hidden
 * `_dsh_体检报告` sheet with score, counts, and per-anomaly rows. The file
 * carries its own audit trail, independent of chat history or external logs.
 * Validators skip `_dsh_` sheets so the report never flags itself.
 */
export declare function writeWorkbookHealthReport(path: string, outPath?: string): Promise<HealthReportResult>;
//# sourceMappingURL=health-report.d.ts.map