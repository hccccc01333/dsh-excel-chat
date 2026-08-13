import { Context } from '@deepseek-ai/cordis'
import { CallId } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import * as plugin from '../src/index.ts'

const ctx = new Context()
await ctx.plugin(SystemPrompt)
await ctx.plugin(ToolRuntime)
await ctx.plugin(plugin)

const schema = ctx.tools.schemas().find((entry) => entry.name === 'excel_validate_formulas')
if (!schema) {
  throw new Error('excel_validate_formulas is not registered')
}
console.log('schema registered:', schema.name)

const result = await ctx.tools.execute({
  signal: new AbortController().signal,
  callId: CallId('vera-p0-1'),
  name: 'excel_validate_formulas',
  arguments: {
    cells: {
      D2: '=B2-C2',
      D3: '=B3-C3',
      D4: '=B4-C3',
      D5: '=B5-C5',
    },
  },
})

console.log('isError:', result.isError)
console.log(JSON.stringify(result.value, null, 2))
await ctx.fiber.dispose()
