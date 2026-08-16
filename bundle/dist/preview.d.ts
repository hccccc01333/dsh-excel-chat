export interface PreviewOptions {
    sheet?: string;
    range?: string;
    maxRows?: number;
}
export interface PreviewResult {
    path: string;
    markdown: string;
    previewPath: string;
    summary: string;
    /** Structured sheet cells so the web panel can render a real spreadsheet grid. */
    sheets: Array<{
        sheet: string;
        cells: Array<{
            id: string;
            value: unknown;
            formula?: string;
            type?: string;
        }>;
    }>;
}
/**
 * Human-facing table preview: render the requested range as a Markdown table
 * (shown inline in the conversation) and write an HTML preview file next to
 * the workbook. Pure ExcelJS + string building, no new dependencies.
 */
export declare function buildWorkbookPreview(path: string, options?: PreviewOptions): Promise<PreviewResult>;
//# sourceMappingURL=preview.d.ts.map