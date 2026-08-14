export interface ReadCell {
    id: string;
    value: string | number | boolean | null;
    formula?: string;
    type: 'string' | 'number' | 'boolean' | 'date' | 'formula' | 'empty';
    numberFormat?: string;
    bold?: boolean;
    italic?: boolean;
    fontSize?: number;
    fontName?: string;
    fontColor?: string;
    fill?: string;
    hAlign?: string;
    vAlign?: string;
    wrapText?: boolean;
    mergedTo?: string;
    dataValidationType?: string;
}
export interface ReadSheetResult {
    sheet: string;
    range: string;
    cells: ReadCell[];
}
export interface ReadWorkbookOptions {
    sheet?: string;
    /** A1 range on the selected sheet, e.g. "A1:D20". */
    range?: string;
    cells?: string[];
}
/** Precisely read cells (values, formulas, types, and formats) from an .xlsx file. */
export declare function readWorkbookDetail(path: string, options?: ReadWorkbookOptions): Promise<ReadSheetResult[]>;
//# sourceMappingURL=read.d.ts.map