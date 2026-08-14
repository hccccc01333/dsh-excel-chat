export type ChartImageExporter = (path: string, outDir: string, signal?: AbortSignal) => Promise<string[]>;
export declare function exportChartsWithExcel(path: string, outDir: string, signal?: AbortSignal): Promise<string[]>;
export type ChartTypeName = 'column' | 'line' | 'pie' | 'bar' | 'area';
export interface ChartCreateOptions {
    sheet?: string;
    /** Data range including headers, e.g. "Sheet1!A1:B4". */
    range: string;
    type?: ChartTypeName;
    title?: string;
    name?: string;
}
export interface ChartModifyOptions {
    type?: ChartTypeName;
    title?: string;
    hasLegend?: boolean;
    axisTitleX?: string;
    axisTitleY?: string;
}
/** Create a chart in an .xlsx copy using local Excel (Windows only). */
export declare function createChartWithExcel(inputPath: string, options: ChartCreateOptions, outPath: string, signal?: AbortSignal): Promise<void>;
/** Modify chart parameters (type, title, legend, axis titles) in an .xlsx copy. */
export declare function modifyChartWithExcel(inputPath: string, chartName: string, changes: ChartModifyOptions, outPath: string, signal?: AbortSignal): Promise<void>;
export type VisionText = (imagePath: string, prompt: string, signal?: AbortSignal) => Promise<string>;
export interface VisualIssue {
    kind: string;
    severity: 'info' | 'warning' | 'critical';
    description: string;
}
export interface ChartVisualReport {
    imagePath: string;
    issues: VisualIssue[];
}
export declare function buildVisionPrompt(): string;
export declare function parseVisionReply(text: string): VisualIssue[];
export declare function createVisionCritic(vision: VisionText): (imagePath: string, signal?: AbortSignal) => Promise<ChartVisualReport>;
export declare function validateChartsVisually(path: string, options: {
    exporter: ChartImageExporter;
    critic: (imagePath: string, signal?: AbortSignal) => Promise<ChartVisualReport>;
    outDir: string;
    signal?: AbortSignal;
}): Promise<{
    images: string[];
    reports: ChartVisualReport[];
}>;
//# sourceMappingURL=chart-visual.d.ts.map