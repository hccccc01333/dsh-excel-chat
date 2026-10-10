/**
 * Analytical findings: what the numbers say, not whether they are clean.
 *
 * `insight.ts` answers "is this data trustworthy" — missing values, duplicates, outliers.
 * It says nothing about what the data *means*, which is the question people actually
 * open a spreadsheet with. Measured before writing any of this: a sheet with a clean
 * 100→260 ramp and one region four times the others produced **zero** findings and no
 * mention of either pattern.
 *
 * Two findings, both deterministic and both stated with the numbers they came from, so
 * a reader can check the claim rather than trust it:
 *
 *   - **trend** — a measure rising or falling across periods of a time column
 *   - **concentration** — a few categories carrying most of a measure
 *
 * No statistics beyond sums and shares. A finding that needs a model to explain itself
 * does not belong in a report a person is expected to act on.
 */
export type AnalysisKind = 'trend' | 'concentration';
export interface AnalysisFinding {
    kind: AnalysisKind;
    /** Where it came from, so the claim can be traced back to columns. */
    sheet: string;
    message: string;
    /** The numbers behind the message. Anything stated is in here. */
    evidence: Record<string, string | number>;
}
/**
 * Analytical findings for a workbook, one sheet at a time.
 *
 * A sheet with no time column still gets concentration; a sheet with no dimension column
 * still gets a trend. A sheet with neither gets nothing — that is the honest answer, not
 * a gap to fill with a weaker claim.
 */
export declare function analyzeWorkbook(path: string, sheet?: string): Promise<AnalysisFinding[]>;
//# sourceMappingURL=analysis.d.ts.map