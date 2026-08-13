import { Context } from '@deepseek-ai/cordis'
import { CallId } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import ExcelJS from 'exceljs'
import { mkdir, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
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
  callId: CallId('vera-repair-1'),
  name: 'excel_repair_formulas',
  arguments: { path },
})

console.log('isError:', result.isError)
const value = result.value as {
  repairs: Array<{ id: string; oldValue: string; newValue: string }>
  before: { anomalies: Array<{ kind: string }> }
  after: { anomalies: Array<{ kind: string }> }
  repairedPath: string
}
console.log('repairs:', JSON.stringify(value.repairs, null, 2))
console.log('before reference-offset:', value.before.anomalies.filter((a) => a.kind === 'reference-offset').length)
console.log('after reference-offset:', value.after.anomalies.filter((a) => a.kind === 'reference-offset').length)
console.log('repairedPath:', value.repairedPath)
