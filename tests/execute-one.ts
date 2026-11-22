import type { Context } from '@deepseek-ai/cordis'
import { ToolCallId } from '@deepseek-ai/dsh-llm'

/**
 * dsh 0.2.2 moved tool execution into explicit executor plugins and changed
 * `ctx.tools.execute` from one call to a raw batch of model calls. Run a single
 * call the way the host does now, and hand back its settled result.
 */
export async function executeOne(ctx: Context, callId: string, name: string, args: unknown) {
  const batch = await ctx.tools.execute({
    signal: new AbortController().signal,
    calls: [{ id: ToolCallId(callId), name, arguments: JSON.stringify(args) }],
  })
  return batch.results[0]!.result
}
