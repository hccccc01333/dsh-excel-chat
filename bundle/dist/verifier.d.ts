/** A deterministic assertion over one workbook cell and its presentation. */
export interface CellAssertion {
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
/**
 * A deterministic assertion over a sheet's view and layout.
 *
 * These operations change nothing a cell assertion can see. Freezing panes, hiding a
 * row, setting a column width — the cell contents are untouched, so the corpus could not
 * assert them at all and 38 of the 77 operations had no end-to-end coverage. This is the
 * missing subject: the sheet itself.
 *
 * Discriminated by `sheet` rather than a `kind` field, because the two variants are
 * already distinguishable and an extra tag would have to be threaded through every
 * existing assertion.
 */
export interface SheetAssertion {
    /** Sheet name. */
    sheet: string;
    /** Required frozen-pane cell (`A2` freezes the first row); `null` requires no freeze. */
    frozen?: string | null;
    /** Rows that must be hidden. */
    hiddenRows?: number[];
    /** Columns that must be hidden. */
    hiddenColumns?: string[];
    /** Required column widths, keyed by column letter. */
    columnWidth?: Record<string, number>;
    /** Required row heights, keyed by row number. */
    rowHeight?: Record<string, number>;
}
export type WorkbookAssertion = CellAssertion | SheetAssertion;
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