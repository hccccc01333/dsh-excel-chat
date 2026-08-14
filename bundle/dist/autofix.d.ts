import { type RepairAdvisor } from './repair.ts';
import type { CellPatch } from './patch.ts';
import type { ValidationResult } from './validator.ts';
export interface AnomalySummary {
    total: number;
    byKind: Record<string, number>;
    formulaCount: number;
    cellCount: number;
}
export interface AutofixOutcome {
    repairedPath: string;
    repairs: CellPatch[];
    before: AnomalySummary;
    after: AnomalySummary;
    message: string;
}
/** Collapse a validation result into counts the model can reason about. */
export declare function summarizeValidation(result: ValidationResult): AnomalySummary;
/**
 * One-call self-healing loop: validate, apply deterministic repairs (plus an
 * optional LLM advisor), re-validate the repaired copy, and report a compact
 * before/after summary in plain language.
 */
export declare function autofixWorkbookFile(path: string, options?: {
    outPath?: string;
    advisor?: RepairAdvisor;
}): Promise<AutofixOutcome>;
//# sourceMappingURL=autofix.d.ts.map