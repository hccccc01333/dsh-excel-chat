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
    /** Required protection state. */
    protected?: boolean;
    /** Required tab colour, with or without an ARGB alpha prefix. */
    tabColor?: string;
    /** Required autofilter range (`A1:B1`); `null` requires no filter. */
    autoFilter?: string | null;
    /** Required cell comments, keyed by cell id, valued by the comment text. */
    comments?: Record<string, string>;
    /** Required hyperlinks, keyed by cell id, valued by the URL. */
    hyperlinks?: Record<string, string>;
    /** Required row outline levels, keyed by row number. */
    outlineRows?: Record<string, number>;
    /** Required column outline levels, keyed by column letter. */
    outlineColumns?: Record<string, number>;
    /** Required number of embedded images. */
    images?: number;
    /** Table names that must exist. */
    tables?: string[];
    /** Required data-validation types, keyed by cell id. */
    dataValidation?: Record<string, string>;
    /** Required page orientation, e.g. `landscape`. */
    pageOrientation?: string;
    /** Required odd-page header string, as written (Excel codes and all). */
    oddHeader?: string;
    /** Required gridline visibility. */
    showGridLines?: boolean;
    /** Required zoom percentage. */
    zoom?: number;
    /** Required sheet state. */
    state?: 'visible' | 'hidden' | 'veryHidden';
    /** Merged ranges that must exist, as written (`A1:B1`). */
    merged?: string[];
}
/**
 * A deterministic assertion over the workbook itself: sheet order and document
 * properties.
 *
 * These belong to no single sheet. `moveSheet` changes the order of the tab strip and
 * `setWorkbookProperties` writes the metadata Excel shows in File → Info; neither is
 * visible from any sheet, which is why neither had coverage.
 */
export interface WorkbookLevelAssertion {
    /** Discriminator. A sheet assertion names a sheet, a cell assertion a cell. */
    workbook: true;
    /** Required sheet order, by name. */
    sheetOrder?: string[];
    /** Required document properties. */
    properties?: {
        title?: string;
        creator?: string;
    };
}
export type WorkbookAssertion = CellAssertion | SheetAssertion | WorkbookLevelAssertion;
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