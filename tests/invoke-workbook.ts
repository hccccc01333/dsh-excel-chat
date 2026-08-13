import { Context } from '@deepseek-ai/cordis'
import { CallId } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import ExcelJS from 'exceljs'
import { fileURLToPath } from 'node:url'
import { mkdir, writeFile } from 'node:fs/promises'
import * as plugin from '../src/index.ts'

const path = fileURLToPath(new URL('../fixtures/silent-error.xlsx', import.meta.url))

const workbook = new ExcelJS.Workbook()
const sheet = workbook.addWorksheet('Sales')
sheet.getCell('D2').value = { formula: 'B2-C2', result: 40 }
sheet.getCell('D3').value = { formula: 'B3-C3', result: 80 }
sheet.getCell('D4').value = { formula: 'B4-C3', result: 150 }
sheet.getCell('D5').value = { formula: 'B5-C5', result: 160 }
await mkdir(fileURLToPath(new URL('../fixtures', import.meta.url)), { recursive: true })
await writeFile(path, await workbook.xlsx.writeBuffer())

const ctx = new Context()
await ctx.plugin(SystemPrompt)
await ctx.plugin(ToolRuntime)
await ctx.plugin(plugin)

const result = await ctx.tools.execute({
  signal: new AbortController().signal,
  callId: CallId('vera-workbook-1'),
  name: 'excel_validate_formulas',
  arguments: { path },
})

console.log('workbook:', path)
console.log('isError:', result.isError)
console.log(JSON.stringify(result.value, null, 2))
await ctx.fiber.dispose()
