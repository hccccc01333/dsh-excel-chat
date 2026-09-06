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
//# sourceMappingURL=plan-schema.d.ts.map