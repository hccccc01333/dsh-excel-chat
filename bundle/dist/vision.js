import { createUserMessage } from '@deepseek-ai/dsh-llm';
import { readFile } from 'node:fs/promises';
/**
 * Wrap the dsh `ctx.llm` streaming service plus `ctx.attachments` into the
 * visual critic's VisionText shape: upload the chart PNG as an image
 * attachment, send it with the checklist prompt, and collect text deltas.
 */
export function visionTextFromContext(ctx, provider, model) {
    return async (imagePath, prompt, signal) => {
        const llm = ctx.get('llm');
        const attachments = ctx.get('attachments');
        if (!llm)
            throw new Error('llm service is not mounted');
        if (!attachments)
            throw new Error('attachments service is not mounted');
        const data = await readFile(imagePath);
        const ref = await attachments.saveImage({
            data,
            mediaType: 'image/png',
            name: imagePath.split(/[\\/]/).pop(),
        });
        const message = createUserMessage({
            content: [
                { type: 'text', text: prompt },
                { type: 'image', attachment: ref },
            ],
            source: { kind: 'user' },
        });
        let text = '';
        for await (const rawChunk of llm.stream({
            provider,
            model,
            messages: [message],
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
