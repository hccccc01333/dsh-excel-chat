import type { ExcelOperation } from './operations.ts';
import type { PlanStep } from './agent.ts';
import type { WorkbookAssertion } from './verifier.ts';
export interface SanitizedPlan {
    steps: PlanStep[];
    notes: string[];
}
/**
 * Normalize planner-produced assertions into WorkbookAssertions (Verifier 2.0):
 * keep only fully-formed entries, coerce ids to strings, and drop style-only
 * assertions that carry no checkable property. Invalid-but-salvageable fields
 * are repaired; entries with neither an id nor any expected value are dropped.
 */
export declare function sanitizeAssertions(assertions: unknown, sheetNames: string[]): {
    assertions: WorkbookAssertion[];
    notes: string[];
};
/**
 * Validate and repair a planner-produced plan before execution. Salvageable
 * issues are fixed in place (sheet prefix, missing sheet, array wrapping,
 * alias fields, cell values); unsalvageable issues throw so the agent loop
 * can feed the exact message back to the planner for a corrected plan.
 */
export declare function sanitizePlan(steps: PlanStep[], sheetNames: string[]): SanitizedPlan;
/**
 * Run the same validation and salvage the planner path uses, for an operations
 * array that arrived straight from a tool call rather than from the planner.
 *
 * `excel_operate` and `excel_workflow` hand the model's array to the executor
 * directly, and the executor assumes it is already well-formed. So a model that
 * omitted a nested field reached a handler and crashed with
 * `TypeError: Cannot read properties of undefined (reading 'toUpperCase')`,
 * where the planner gets the actionable `sortRange 的 keys[0].column 缺失`.
 * Reusing this function rather than writing a second validator is deliberate:
 * one rule with two implementations drifts.
 */
export declare function sanitizeOperations(operations: ExcelOperation[], sheetNames: string[]): ExcelOperation[];
//# sourceMappingURL=plan-schema.d.ts.map