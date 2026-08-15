import type { LlmText } from './advisor.ts';
import type { AgentPlanner } from './agent.ts';
/**
 * LLM-backed planner/verifier for the goal-driven agent loop. The planner
 * proposes excel_operate steps from the goal + file profile; the verifier
 * judges whether the executed result achieved the goal.
 */
export declare function createLlmPlanner(llm: LlmText): AgentPlanner;
//# sourceMappingURL=llm-planner.d.ts.map