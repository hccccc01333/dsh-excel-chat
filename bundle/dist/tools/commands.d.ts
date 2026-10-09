/**
 * The three slash commands.
 *
 * They act on a local file in place rather than producing a new one, which is why
 * they are commands and not tools: the user typed a path and expects that file to
 * change. Split out of `index.ts` unchanged.
 */
import type { Context } from '@deepseek-ai/cordis';
export declare function registerCommands(ctx: Context): void;
//# sourceMappingURL=commands.d.ts.map