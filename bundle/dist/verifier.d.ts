/** A deterministic assertion over one workbook cell and its presentation. */
export interface WorkbookAssertion {
    /** Cell id, such as `订单!B3`. */
    id: string;
    /** Exact serialized cell content; `null` requires an empty or absent cell. */
    expect?: string | null;
    /** Required serialized cell-content prefix, useful for formulas. */
    startsWith?: string;
    /** Required foreground fill color, with or without an ARGB alpha prefix. */
    fill?: string;
    /** Required bold state. */
    bold?: boolean;
    /** Required Excel number format. */
    numberFormat?: string;
    /** Required wrap-text state. */
    wrapText?: boolean;
    /** Required horizontal alignment. */
    hAlign?: string;
}
/** The result for one deterministic workbook assertion. */
export interface WorkbookAssertionResult {
    id: string;
    passed: boolean;
    detail: string;
}
/** Evidence returned by the deterministic verifier. */
export interface WorkbookVerification {
    achieved: boolean;
    passed: number;
    total: number;
    failures: string[];
    assertions: WorkbookAssertionResult[];
    reason: string;
}
/**
 * Evaluate workbook assertions without an LLM.
 *
 * Cell-content assertions use the normalized workbook cell map. Presentation
 * assertions load the workbook once and require every declared presentation
 * property on a check to match.
 */
export declare function verifyWorkbookAssertions(path: string, assertions: WorkbookAssertion[]): Promise<WorkbookVerification>;
//# sourceMappingURL=verifier.d.ts.map