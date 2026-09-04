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
export interface WorkbookAnnotations {
    comments: Map<string, CommentSpec[]>;
    sparklines: Map<string, SparklineGroupSpec[]>;
}
export declare function emptyAnnotations(): WorkbookAnnotations;
/**
 * Rewrite the saved xlsx zip: inject comments parts, VML shapes, sparkline
 * extensions, and the worksheet plumbing (legacyDrawing + rels + content
 * types) they require.
 */
export declare function annotateWorkbookXml(data: Uint8Array, annotations: WorkbookAnnotations, sheetFileOf: Map<string, string>): Uint8Array;
//# sourceMappingURL=xml-postprocess.d.ts.map