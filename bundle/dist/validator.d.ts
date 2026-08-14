import { type DependencyGraph } from './graph.ts';
import { type ColumnPatternReport, type PatternAnomaly } from './patterns.ts';
export interface ValidationResult {
    cellCount: number;
    formulaCount: number;
    dependencyGraph: DependencyGraph;
    columns: ColumnPatternReport[];
    anomalies: PatternAnomaly[];
    errors: Array<{
        id: string;
        message: string;
    }>;
}
export declare function validate(cells: Record<string, string>): ValidationResult;
//# sourceMappingURL=validator.d.ts.map