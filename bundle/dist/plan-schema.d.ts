import type { PlanStep } from './agent.ts';
export interface SanitizedPlan {
    steps: PlanStep[];
    notes: string[];
}
/**
 * Validate and repair a planner-produced plan before execution. Salvageable
 * issues are fixed in place (sheet prefix, missing sheet, array wrapping,
 * alias fields, cell values); unsalvageable issues throw so the agent loop
 * can feed the exact message back to the planner for a corrected plan.
 */
export declare function sanitizePlan(steps: PlanStep[], sheetNames: string[]): SanitizedPlan;
//# sourceMappingURL=plan-schema.d.ts.map