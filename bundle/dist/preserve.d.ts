export interface PreservedContent {
    /** Parts to restore verbatim, keyed by their path inside the package. */
    parts: Record<string, Uint8Array>;
    /** `<pivotCaches>` from the original workbook.xml, if any. */
    pivotCaches?: string;
    /** The `pivotCacheDefinition` relationship from the original workbook rels. */
    cacheRelationship?: string;
    /** Sheet name → `pivotTable` relationships declared for that sheet. */
    sheetRelationships: Map<string, string[]>;
    /** Sheet name → the `<tableParts>` element the worksheet carried, if any. */
    sheetTableParts: Map<string, string>;
    /** `<Override>` entries covering the preserved parts. */
    contentTypes: string[];
}
/** True when there is nothing worth carrying through an edit. */
export declare function isEmpty(preserved: PreservedContent): boolean;
/** Read everything an edit would otherwise destroy out of the original file. */
export declare function capturePreserved(data: Uint8Array): PreservedContent;
/**
 * Put the captured content back into a freshly written package.
 *
 * Exported for testing: it is pure, so the merge can be checked against real
 * packages without going anywhere near a file.
 */
export declare function applyPreserved(output: Uint8Array, preserved: PreservedContent): Uint8Array;
//# sourceMappingURL=preserve.d.ts.map