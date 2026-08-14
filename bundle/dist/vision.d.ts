import type { Context } from '@deepseek-ai/cordis';
import type { VisionText } from './chart-visual.ts';
/**
 * Wrap the dsh `ctx.llm` streaming service plus `ctx.attachments` into the
 * visual critic's VisionText shape: upload the chart PNG as an image
 * attachment, send it with the checklist prompt, and collect text deltas.
 */
export declare function visionTextFromContext(ctx: Context, provider: string, model: string): VisionText;
//# sourceMappingURL=vision.d.ts.map