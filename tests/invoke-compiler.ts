import { Context } from '@deepseek-ai/cordis'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import * as NativeExecutor from '@deepseek-ai/dsh-tool-executor-native'
import * as plugin from '../src/index.ts'
import { executeOne } from './execute-one.ts'

const ctx = new Context()
await ctx.plugin(SystemPrompt)
await ctx.plugin(ToolRuntime)
await ctx.plugin(NativeExecutor)
await ctx.plugin(plugin)

const result = await executeOne(ctx, 'vera-compiler-1', 'excel_compile_formula', {
  ir: {
    operation: 'aggregate',
    metric: 'sales',
    function: 'SUMIFS',
    filters: [{ column: 'channel', value_from: 'A2' }],
  },
  baseCell: 'B2',
  table: { sheet: 'Sales', columns: { sales: 'H', channel: 'C' } },
})

console.log('isError:', result.isError)
console.log(JSON.stringify(result.value, null, 2))
