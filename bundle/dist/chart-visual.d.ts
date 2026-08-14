export type ChartImageExporter = (path: string, outDir: string, signal?: AbortSignal) => Promise<string[]>;
export declare function exportChartsWithExcel(path: string, outDir: string, signal?: AbortSignal): Promise<string[]>;
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