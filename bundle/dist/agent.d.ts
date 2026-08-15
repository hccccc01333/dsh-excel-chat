import type { ExcelOperation } from './operations.ts';
import { type TaskResult } from './task.ts';
export interface PlanStep {
    name?: string;
    operations: ExcelOperation[];
}
export interface AgentPlanContext {
    goal: string;
    path: string;
    round: number;
    profileSummary: string;
    validationSummary: string;
    previousPlan?: PlanStep[];
    previousResult?: TaskResult;
    verifierNote?: string;
}
export interface AgentVerifierContext extends AgentPlanContext {
    /** Plan that was just executed; its output is at `path`. */
    executedPlan: PlanStep[];
    executedResult: TaskResult;
}
export interface AgentPlanner {
    plan(context: AgentPlanContext): Promise<PlanStep[]>;
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
}): Promise<AgentTaskResult>;
//# sourceMappingURL=agent.d.ts.map