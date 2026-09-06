import type { ExcelOperation } from './operations.ts';
import { type TaskResult } from './task.ts';
import { type WorkbookAssertion, type WorkbookVerification } from './verifier.ts';
export interface PlanStep {
    name?: string;
    operations: ExcelOperation[];
}
/**
 * Verifier 2.0: the planner may accompany its steps with machine-checkable
 * assertions ("expect") that are verified deterministically against the
 * executed workbook, in addition to the LLM verdict.
 */
export interface PlanWithAssertions {
    steps: PlanStep[];
    assertions?: WorkbookAssertion[];
}
export type PlannerPlanOutput = PlanStep[] | PlanWithAssertions;
export interface AgentPlanContext {
    goal: string;
    path: string;
    round: number;
    sheetNames: string[];
    profileSummary: string;
    semanticSummary?: string;
    validationSummary: string;
    previousPlan?: PlanStep[];
    previousResult?: TaskResult;
    verifierNote?: string;
}
export interface AgentVerifierContext extends AgentPlanContext {
    /** Plan that was just executed; its output is at `path`. */
    executedPlan: PlanStep[];
    executedResult: TaskResult;
    /** Compact post-execution cell snapshot for evidence-based verification. */
    cellSnapshot: string;
}
export interface AgentPlanner {
    plan(context: AgentPlanContext): Promise<PlannerPlanOutput>;
    verify(context: AgentVerifierContext): Promise<{
        achieved: boolean;
        reason: string;
    }>;
}
export interface AgentRoundResult {
    round: number;
    plan: PlanStep[];
    result: TaskResult;
    verdict: {
        achieved: boolean;
        reason: string;
    };
    deterministicVerification?: WorkbookVerification;
    /** Verifier 2.0: deterministic check of the planner's own assertions. */
    planAssertions?: WorkbookVerification;
}
export interface AgentTaskResult {
    outputPath: string;
    rounds: AgentRoundResult[];
    achieved: boolean;
    finalAnomalies: number;
}
/**
 * Goal-driven agent loop (Plan -> Act -> Observe -> Verify -> Replan):
 * the planner proposes operation steps for the goal, `runExcelTask` executes
 * them with per-step formula verification and deterministic repair, an LLM
 * verifier checks whether the goal is achieved, and the loop replans up to
 * maxRounds times when it is not.
 */
export declare function runAgentTask(path: string, options: {
    goal: string;
    planner: AgentPlanner;
    maxRounds?: number;
    outPath?: string;
    /** Optional hard assertions used by deterministic benchmark/replay callers. */
    deterministicAssertions?: WorkbookAssertion[];
}): Promise<AgentTaskResult>;
//# sourceMappingURL=agent.d.ts.map