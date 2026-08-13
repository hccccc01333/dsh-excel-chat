import { createUserMessage } from '@deepseek-ai/dsh-llm';
/**
 * Wrap the dsh `ctx.llm` streaming service into the advisor's LlmText shape.
 * Throws when the llm service is not mounted or the stream finishes with an
 * error/aborted reason.
 */
export function llmTextFromContext(ctx, provider, model) {
    return async (prompt, signal) => {
        const llm = ctx.get('llm');
        if (!llm)
            throw new Error('llm service is not mounted');
        let text = '';
        for await (const rawChunk of llm.stream({
            provider,
            model,
            messages: [createUserMessage({
                    content: [{ type: 'text', text: prompt }],
                    source: { kind: 'user' },
                })],
            temperature: 0,
            maxTokens: 2000,
            signal,
        })) {
            const chunk = rawChunk;
            if (chunk.type === 'text-delta' && typeof chunk.text === 'string') {
                text += chunk.text;
            }
            else if (chunk.type === 'finish') {
                if (chunk.reason?.kind === 'error' || chunk.reason?.kind === 'aborted') {
                    throw new Error(`LLM call failed: ${chunk.reason.failure?.message ?? chunk.reason.kind}`);
                }
            }
        }
        return text;
    };
}
