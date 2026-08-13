import { existsSync } from 'node:fs'
import { mkdir, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import { createLlmRepairAdvisor } from '../src/advisor.ts'
import { deepseekLlmTextFromEnv } from '../src/deepseek.ts'
import { repairWorkbookFile } from '../src/repair.ts'

if (!process.env.DEEPSEEK_API_KEY) {
  console.log('DEEPSEEK_API_KEY is not set; skipping real LLM invocation')
  process.exit(0)
}

const path = fileURLToPath(new URL('../fixtures/structure-mismatch.xlsx', import.meta.url))
const workbook = new ExcelJS.Workbook()
const sheet = workbook.addWorksheet('Sheet1')
sheet.getCell('D2').value = { formula: 'B2-C2', result: 40 }
sheet.getCell('D3').value = { formula: 'SUM(B3:C3)', result: 80 }
sheet.getCell('D4').value = { formula: 'B4-C4', result: 80 }
await mkdir(fileURLToPath(new URL('../fixtures', import.meta.url)), { recursive: true })
await writeFile(path, await workbook.xlsx.writeBuffer())

const table = { sheet: 'Sheet1', columns: { revenue: 'B', cost: 'C' } }
const advisor = createLlmRepairAdvisor(deepseekLlmTextFromEnv(), table)
const result = await repairWorkbookFile(path, advisor)

console.log('model:', process.env.DEEPSEEK_MODEL ?? 'deepseek-chat')
console.log('llmRepairs:', JSON.stringify(result.llmRepairs, null, 2))
console.log('before structure-mismatch:', result.before.anomalies.filter((a) => a.kind === 'structure-mismatch').length)
console.log('after structure-mismatch:', result.after.anomalies.filter((a) => a.kind === 'structure-mismatch').length)
console.log('repairedPath:', result.repairedPath)
if (!existsSync(result.repairedPath)) {
  throw new Error('repaired file was not written')
}
