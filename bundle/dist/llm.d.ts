import type { Context } from '@deepseek-ai/cordis';
import type { LlmText } from './advisor.ts';
/**
 * Wrap the dsh `ctx.llm` streaming service into the advisor's LlmText shape.
 * Throws when the llm service is not mounted or the stream finishes with an
 * error/aborted reason.
 */
export declare function llmTextFromContext(ctx: Context, provider: string, model: string): LlmText;
//# sourceMappingURL=llm.d.ts.map