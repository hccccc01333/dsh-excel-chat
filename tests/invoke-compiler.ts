import { Context } from '@deepseek-ai/cordis'
import { CallId } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import * as plugin from '../src/index.ts'

const ctx = new Context()
await ctx.plugin(SystemPrompt)
await ctx.plugin(ToolRuntime)
await ctx.plugin(plugin)

const result = await ctx.tools.execute({
  signal: new AbortController().signal,
  callId: CallId('vera-compiler-1'),
  name: 'excel_compile_formula',
  arguments: {
    ir: {
      operation: 'aggregate',
      metric: 'sales',
      function: 'SUMIFS',
      filters: [{ column: 'channel', value_from: 'A2' }],
    },
    baseCell: 'B2',
    table: { sheet: 'Sales', columns: { sales: 'H', channel: 'C' } },
  },
})

console.log('isError:', result.isError)
console.log(JSON.stringify(result.value, null, 2))
