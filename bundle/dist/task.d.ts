import { type ExcelOperation, type OperationWarning } from './operations.ts';
export interface TaskStepInput {
    name?: string;
    operations: ExcelOperation[];
    /** Validate formulas after this step and auto-repair anomalies (default true). */
    verify?: boolean;
}
export interface TaskStepResult {
    index: number;
    name: string;
    warnings: OperationWarning[];
    validation?: {
        before: number;
        after: number;
        fixed: number;
    };
}
export interface TaskResult {
    outputPath: string;
    steps: TaskStepResult[];
    finalAnomalies: number;
}
/**
 * Multi-step Excel workflow in one call (ExcelGenius2/SheetCopilot-style task
 * orchestration): each step applies an operations array, validates formulas,
 * auto-repairs anomalies with the deterministic fixer, and feeds the verified
 * result into the next step. Intermediate files stay in a temp dir; only the
 * final output is copied to outPath.
 */
export declare function runExcelTask(path: string, steps: TaskStepInput[], outPath?: string): Promise<TaskResult>;
//# sourceMappingURL=task.d.ts.map