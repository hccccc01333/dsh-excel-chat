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

const schema = ctx.tools.schemas().find((entry) => entry.name === 'excel_validate_formulas')
if (!schema) {
  throw new Error('excel_validate_formulas is not registered')
}
console.log('schema registered:', schema.name)

const result = await executeOne(ctx, 'vera-p0-1', 'excel_validate_formulas', {
  cells: {
    D2: '=B2-C2',
    D3: '=B3-C3',
    D4: '=B4-C3',
    D5: '=B5-C5',
  },
})

console.log('isError:', result.isError)
console.log(JSON.stringify(result.value, null, 2))
await ctx.fiber.dispose()
