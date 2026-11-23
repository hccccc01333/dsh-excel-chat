/**
 * Post-save xlsx XML annotations for features ExcelJS cannot write:
 * cell comments (comments XML + VML shapes), and sparklines (x14 ext).
 *
 * ExcelJS reads comments but silently drops them on save, so comments added
 * by `addComment` are re-injected into the saved zip here. Sparklines are a
 * Microsoft x14 worksheet extension that ExcelJS never models.
 */
export interface CommentSpec {
    /** 1-based cell reference inside the sheet, e.g. "B2". */
    ref: string;
    text: string;
    author: string;
    /** Comment box size in points (defaults 108 x 60). */
    width: number;
    height: number;
}
export interface SparklineGroupSpec {
    /** Workbook-qualified data range, e.g. "订单!B2:F31". */
    dataRange: string;
    /** Workbook-qualified location range, e.g. "订单!G2:G31". */
    locationRange: string;
    type: 'line' | 'column' | 'stacked';
    color: string;
    negativeColor: string;
    markers: boolean;
    highColor: string;
    lowColor: string;
}
/** A saved autoFilter criterion ExcelJS cannot express on its own. */
export interface AutoFilterColumnSpec {
    /** Zero-based column index within the autoFilter range. */
    colId: number;
    /** Exact values to select (the <filters> form). */
    values?: string[];
    /** Comparison to write as <customFilters>, e.g. { operator: 'greaterThan', val: '1000' }. */
    custom?: {
        operator: string;
        val: string;
    };
}
export interface ChartSeries {
    name: string;
    ref: string;
    values: number[];
}
/** A chart ExcelJS cannot create, described entirely in worksheet coordinates. */
export interface ChartSpec {
    type: 'column' | 'bar' | 'line' | 'pie';
    title?: string;
    categoryRef: string;
    categories: string[];
    series: ChartSeries[];
    legend?: boolean;
    xAxisTitle?: string;
    yAxisTitle?: string;
    /** Zero-based top-left cell for the chart. */
    anchorCol: number;
    anchorRow: number;
    /** Drawing size in EMU (914400 per inch). */
    width: number;
    height: number;
}
export type PivotAggregate = 'sum' | 'average' | 'count' | 'max' | 'min';
export interface PivotCacheField {
    name: string;
    /** Numeric fields carry raw values in records; text fields carry shared-item indexes. */
    numeric: boolean;
    /** Shared string items for text fields. */
    items?: string[];
    min?: number;
    max?: number;
}
export interface PivotMeasure {
    name: string;
    /** Zero-based index into PivotSpec.fields. */
    fieldIndex: number;
    aggregate: PivotAggregate;
    /** Aggregated values in row order. */
    values: number[];
    /** Row-by-column values when a column field is present. */
    matrix?: number[][];
}
/** A native pivot table built from a workbook range, without Excel. */
export interface PivotSpec {
    name: string;
    cacheId: number;
    /** Part-number suffix, kept clear of pivot parts restored from the original file. */
    partSuffix: number;
    sourceSheet: string;
    /** Source range body without the sheet name, e.g. "A1:B4". */
    sourceRef: string;
    /** Destination sheet and its body range, e.g. "透视!A3:B6". */
    outputSheet: string;
    outputRef: string;
    fields: PivotCacheField[];
    /** Zero-based index of the row field in `fields`. */
    rowFieldIndex: number;
    /** Shared-item index of each output row's row-field value. */
    rowItemIndices: number[];
    /** Zero-based index of the optional column field in `fields`. */
    columnFieldIndex?: number;
    /** Shared-item index of each output column's column-field value. */
    columnItemIndices?: number[];
    /** Zero-based index of the optional page (filter) field in `fields`. */
    pageFieldIndex?: number;
    /** Shared-item index selected by the page field. */
    pageItemIndex?: number;
    measures: PivotMeasure[];
    /** One entry per source data row, one number per field. */
    records: number[][];
}
export interface WorkbookAnnotations {
    comments: Map<string, CommentSpec[]>;
    sparklines: Map<string, SparklineGroupSpec[]>;
    /** Sheets whose saved view should show formulas instead of their results. */
    showFormulas: Set<string>;
    /** Criteria to inject into an autoFilter that ExcelJS wrote as a bare ref. */
    filterColumns: Map<string, AutoFilterColumnSpec[]>;
    /** Charts to add as chart/drawing parts. */
    charts: Map<string, ChartSpec[]>;
    /** Native pivot tables to add as pivotCache/pivotTable parts. */
    pivots: PivotSpec[];
}
export declare function emptyAnnotations(): WorkbookAnnotations;
/**
 * Rewrite the saved xlsx zip: inject comments parts, VML shapes, sparkline
 * extensions, and the worksheet plumbing (legacyDrawing + rels + content
 * types) they require.
 */
export declare function annotateWorkbookXml(data: Uint8Array, annotations: WorkbookAnnotations, sheetFileOf: Map<string, string>): Uint8Array;
//# sourceMappingURL=xml-postprocess.d.ts.map