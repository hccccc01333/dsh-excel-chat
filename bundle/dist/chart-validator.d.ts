import type { ChartInfo } from './charts.ts';
export type ChartAnomalyKind = 'unknown-chart-type' | 'no-series' | 'missing-categories' | 'missing-values' | 'invalid-range' | 'missing-cells' | 'multi-dimensional-range' | 'unsorted-dates';
export interface ChartAnomaly {
    kind: ChartAnomalyKind;
    chartPath: string;
    seriesIndex: number | null;
    message: string;
}
export interface ChartValidationReport {
    sheetName: string;
    chartPath: string;
    type: string | null;
    seriesCount: number;
    anomalies: ChartAnomaly[];
}
export declare function validateCharts(charts: ChartInfo[], cells: Record<string, string>): ChartValidationReport[];
//# sourceMappingURL=chart-validator.d.ts.map