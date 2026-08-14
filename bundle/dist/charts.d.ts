export interface ChartSeries {
    name?: string;
    categories?: string;
    values?: string;
}
export interface ChartInfo {
    sheetName: string;
    chartPath: string;
    type: string | null;
    series: ChartSeries[];
}
export declare function readChartInfos(path: string): Promise<ChartInfo[]>;
export declare function parseChartXml(xml: string): {
    type: string | null;
    series: ChartSeries[];
};
export interface ResolvedRange {
    sheet: string;
    startColumn: string;
    startRow: number;
    endColumn: string;
    endRow: number;
}
export declare function parseRangeRef(ref: string): ResolvedRange | null;
export declare function expandRange(range: ResolvedRange): string[];
//# sourceMappingURL=charts.d.ts.map