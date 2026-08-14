import type { ColumnTable, FormulaIR } from './ir.ts';
import type { CellPatch } from './patch.ts';
import type { ColumnPatternReport, PatternAnomaly } from './patterns.ts';
import type { ValidationResult } from './validator.ts';
export type LlmText = (prompt: string, signal?: AbortSignal) => Promise<string>;
export interface LlmRepairReply {
    repairs: Array<{
        id: string;
        baseCell: string;
        ir: FormulaIR;
    }>;
}
export declare function buildRepairPrompt(cells: Record<string, string>, anomalies: PatternAnomaly[], table: ColumnTable, columns?: ColumnPatternReport[]): string;
/**
 * Wrap an LLM text function into a repair advisor: given the workbook excerpt,
 * validation anomalies, and the table schema, ask the model for IR repairs and
 * compile them into concrete CellPatches.
 */
export declare function createLlmRepairAdvisor(llm: LlmText, table: ColumnTable, signal?: AbortSignal): (cells: Record<string, string>, result: ValidationResult) => Promise<CellPatch[]>;
/**
 * Tolerate common model mistakes: bare strings as operands become cell or
 * column operands instead of failing schema validation, and an aggregate SUM
 * without a metric falls back to the table's first column.
 */
export declare function normalizeIr(ir: FormulaIR, table?: ColumnTable): FormulaIR;
//# sourceMappingURL=advisor.d.ts.map