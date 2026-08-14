import ExcelJS from 'exceljs'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Context } from '@deepseek-ai/cordis'
import { CallId } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import { exportChartsWithExcel } from '../src/chart-visual.ts'
import * as plugin from '../src/index.ts'
import { deepseekChatWithTools, type DeepSeekMessage } from '../src/deepseek.ts'
import { readWorkbookCells } from '../src/workbook.ts'
import { makeSalesWorkbook } from './sales-fixture.ts'

if (!process.env.DEEPSEEK_API_KEY) {
  console.log('DEEPSEEK_API_KEY is not set; skipping real conversation invocation')
  process.exit(0)
}

const apiKey = process.env.DEEPSEEK_API_KEY
const baseUrl = process.env.DEEPSEEK_BASE_URL
const model = process.env.DEEPSEEK_MODEL ?? 'deepseek-chat'
const dir = await mkdtemp(join(tmpdir(), 'vera-conversation-'))

async function writeFixture(name: string, build: (sheet: ExcelJS.Worksheet) => void): Promise<string> {
  const path = join(dir, name)
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  build(sheet)
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

async function readSheet(path: string): Promise<ExcelJS.Worksheet> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(path)
  return workbook.getWorksheet('Sheet1')!
}

async function readSheetByName(path: string, name: string): Promise<ExcelJS.Worksheet> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(path)
  const sheet = workbook.getWorksheet(name)
  if (!sheet) throw new Error(`sheet not found: ${name}`)
  return sheet
}

async function runConversation(
  ctx: Context,
  tools: unknown[],
  request: string,
  maxTurns = 6,
): Promise<number> {
  const messages: DeepSeekMessage[] = [
    {
      role: 'system',
      content:
        'You are an Excel automation agent inside DeepSeek Harness. Use the provided excel_* tools. ' +
        'File paths are absolute. Prefer one excel_operate call with multiple operations for editing. ' +
        'excel_operate writes a new .xlsx and returns post-operation validation; read that output when you need to verify.',
    },
    { role: 'user', content: request },
  ]
  for (let turn = 0; turn < maxTurns; turn++) {
    const reply = await deepseekChatWithTools({
      apiKey,
      baseUrl,
      model,
      messages,
      tools,
      temperature: 0,
      maxTokens: 4000,
    })
    console.log(`--- turn ${turn}: toolCalls=${reply.toolCalls.length} content=${(reply.content ?? '').slice(0, 200)}`)
    if (reply.toolCalls.length === 0) {
      return turn + 1
    }
    messages.push({ role: 'assistant', content: reply.content, tool_calls: reply.toolCalls })
    for (const call of reply.toolCalls) {
      let args: unknown
      try {
        args = JSON.parse(call.function.arguments)
      } catch {
        args = { _parseError: `invalid JSON: ${call.function.arguments}` }
      }
      console.log(`tool: ${call.function.name}`, JSON.stringify(args).slice(0, 500))
      const result = await ctx.tools.execute({
        signal: new AbortController().signal,
        callId: CallId(`vera-conversation-${Date.now()}-${call.id}`),
        name: call.function.name,
        arguments: args ?? {},
      })
      const content = JSON.stringify(result.value ?? (result.isError ? { isError: true } : {}))
      messages.push({ role: 'tool', tool_call_id: call.id, content: content.slice(0, 8000) })
      console.log(`  -> isError=${result.isError}`)
    }
  }
  return maxTurns
}

const ctx = new Context()
await ctx.plugin(SystemPrompt)
await ctx.plugin(ToolRuntime)
await ctx.plugin(plugin)
const tools = ctx.tools.schemas().map((schema) => ({
  type: 'function',
  function: { name: schema.name, description: schema.description, parameters: schema.parameters },
}))
console.log('registered tools:', ctx.tools.schemas().map((schema) => schema.name).join(', '))

// Scenario 1: build a styled report with formulas, freeze panes, and a filter.
const reportPath = await writeFixture('report.xlsx', (sheet) => {
  sheet.getCell('A1').value = '月份'
  sheet.getCell('B1').value = '收入'
  sheet.getCell('C1').value = '成本'
  sheet.getCell('D1').value = '毛利'
  for (let row = 2; row <= 5; row++) {
    sheet.getCell(`A${row}`).value = row - 1
    sheet.getCell(`B${row}`).value = 100 + row
    sheet.getCell(`C${row}`).value = 60 + row
    sheet.getCell(`D${row}`).value = { formula: `B${row}-C${row}` }
  }
})
const reportOut = join(dir, 'report-final.xlsx')
const reportTurns = await runConversation(
  ctx,
  tools,
  `请对文件 ${reportPath} 完成以下操作：1) 保持 D2:D5 的毛利公式 =B行-C行；` +
    `2) 新增 E 列表头“合计”，E2:E5 写公式 =B行+C行；3) 表头 A1:E1 加粗并填充浅灰色；` +
    `4) 冻结第一行；5) 给 A1:E5 加自动筛选。输出到 ${reportOut}。`,
)
const reportSheet = await readSheet(reportOut)
const reportChecks = [
  ['E1 header', reportSheet.getCell('E1').value, '合计'],
  ['E2 formula', reportSheet.getCell('E2').formula, 'B2+C2'],
  ['A1 bold', reportSheet.getCell('A1').font?.bold, true],
  ['freeze panes', reportSheet.views[0]?.state, 'frozen'],
  ['autoFilter', reportSheet.autoFilter, 'A1:E5'],
]
for (const [label, actual, expected] of reportChecks) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected)
  console.log(`[${ok ? 'PASS' : 'FAIL'}] ${label}: ${JSON.stringify(actual)}`)
  if (!ok) throw new Error(`report scenario failed: ${label}`)
}

// Scenario 2: detect and repair a silent formula error through conversation.
const brokenPath = await writeFixture('broken.xlsx', (sheet) => {
  sheet.getCell('B2').value = 10
  sheet.getCell('C2').value = 4
  sheet.getCell('D2').value = { formula: 'B2-C2' }
  sheet.getCell('B3').value = 20
  sheet.getCell('C3').value = 5
  sheet.getCell('D3').value = { formula: 'B3-C3' }
  sheet.getCell('B4').value = 30
  sheet.getCell('C4').value = 6
  sheet.getCell('D4').value = { formula: 'B4-C3' }
  sheet.getCell('B5').value = 40
  sheet.getCell('C5').value = 7
  sheet.getCell('D5').value = { formula: 'B5-C5' }
})
const brokenOut = join(dir, 'broken-repaired.xlsx')
const repairTurns = await runConversation(
  ctx,
  tools,
  `检查文件 ${brokenPath} 里 D 列的公式是否一致，D 列每一行都应该是 =B行-C行。` +
    `如果发现异常，用 excel_repair_formulas 修复（不需要 LLM），输出到 ${brokenOut}。`,
)
const repairedSheet = await readSheet(brokenOut)
const repairChecks = [
  ['D4 repaired', repairedSheet.getCell('D4').formula, 'B4-C4'],
  ['D2 intact', repairedSheet.getCell('D2').formula, 'B2-C2'],
]
for (const [label, actual, expected] of repairChecks) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected)
  console.log(`[${ok ? 'PASS' : 'FAIL'}] ${label}: ${JSON.stringify(actual)}`)
  if (!ok) throw new Error(`repair scenario failed: ${label}`)
}

// Scenario 3: analysis report — pivot-style summary, subtotals, data bars, protection.
const salesPath = await makeSalesWorkbook()
const analysisOut = join(dir, '分析报表.xlsx')
const analysisTurns = await runConversation(
  ctx,
  tools,
  `请对文件 ${salesPath} 做数据分析：` +
    `1) 用 excel_operate 的 aggregateReport 按“区域”生成透视汇总表（金额合计 + 数量计数），输出到工作表“订单-汇总”；` +
    `2) 给“订单”表按“区域”做分类汇总（金额小计 + 总计）；` +
    `3) 给“订单”表 F 列金额加数据条条件格式；` +
    `4) 最后保护“订单”工作表（密码 pw）。输出到 ${analysisOut}。`,
)
const summarySheet = await readSheetByName(analysisOut, '订单-汇总')
const analysisChecks = [
  ['summary SUMIFS', summarySheet.getCell('B2').formula?.startsWith('SUMIFS'), true],
  ['summary group', summarySheet.getCell('A2').value, '华东'],
]
for (const [label, actual, expected] of analysisChecks) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected)
  console.log(`[${ok ? 'PASS' : 'FAIL'}] ${label}: ${JSON.stringify(actual)}`)
  if (!ok) throw new Error(`analysis scenario failed: ${label}`)
}
const orderSheet = await readSheetByName(analysisOut, '订单')
const orderCells = await readWorkbookCells(await readFile(analysisOut))
const hasSubtotalRow = Object.values(orderCells).some((value) => typeof value === 'string' && value.includes('汇总'))
const hasSubtotalFormula = Object.values(orderCells).some((value) => typeof value === 'string' && value.startsWith('=SUBTOTAL'))
const cfTypes = (orderSheet as unknown as { conditionalFormattings?: Array<{ rules: Array<{ type: string }> }> }).conditionalFormattings
  ?.flatMap((entry) => entry.rules.map((rule) => rule.type)) ?? []
console.log(`[${hasSubtotalRow ? 'PASS' : 'FAIL'}] subtotal row present`)
console.log(`[${hasSubtotalFormula ? 'PASS' : 'FAIL'}] SUBTOTAL formula present`)
console.log(`[${cfTypes.includes('dataBar') ? 'PASS' : 'FAIL'}] dataBar conditional formatting present`)
if (!hasSubtotalRow || !hasSubtotalFormula || !cfTypes.includes('dataBar')) {
  throw new Error('analysis scenario failed: missing subtotals or data bars')
}

// Scenario 4: VLOOKUP enrichment + mail merge through conversation.
const mergeOut = join(dir, 'vlookup-merge.xlsx')
const mergeTurns = await runConversation(
  ctx,
  tools,
  `请处理文件 ${salesPath}：` +
    `1) 给“订单”表新增 G 列“产品名称”，用 VLOOKUP 从“产品价目”表（A1:B4）按产品代码查找名称，G2:G7 各一行；` +
    `2) 用 excel_operate 的 mailMerge，以“通知模板”为模板、订单表 A:D 为数据，批量生成发货通知到工作表“发货通知”。` +
    `输出到 ${mergeOut}。`,
)
const mergeOrders = await readSheetByName(mergeOut, '订单')
const noticeSheet = await readSheetByName(mergeOut, '发货通知')
const mergeChecks = [
  ['VLOOKUP header', mergeOrders.getCell('G1').value, '产品名称'],
  ['VLOOKUP formula', mergeOrders.getCell('G2').formula?.startsWith('VLOOKUP(C2,'), true],
  ['merge row 1', noticeSheet.getCell('A1').value, '华东'],
  ['merge row 1 qty', noticeSheet.getCell('B1').value, '数量 10'],
  ['merge rows', noticeSheet.getCell('A6').value, '华东'],
]
for (const [label, actual, expected] of mergeChecks) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected)
  console.log(`[${ok ? 'PASS' : 'FAIL'}] ${label}: ${JSON.stringify(actual)}`)
  if (!ok) throw new Error(`merge scenario failed: ${label}`)
}

// Scenario 5: create a chart through conversation (Windows + Excel only).
let chartTurns = 0
if (process.platform === 'win32') {
  const chartOut = join(dir, '图表.xlsx')
  chartTurns = await runConversation(
    ctx,
    tools,
    `请用 excel_create_chart 给文件 ${salesPath} 的“订单”表创建柱状图：` +
      `数据范围 “订单!B1:F7”（区域为分类、金额为数值），标题“区域金额”，图表名“Chart 1”。输出到 ${chartOut}。`,
  )
  const images = await exportChartsWithExcel(chartOut, join(dir, 'chart-out'))
  console.log(`[${images.length >= 1 ? 'PASS' : 'FAIL'}] chart created and exported (${images.length} chart(s))`)
  if (images.length < 1) throw new Error('chart scenario failed')
} else {
  console.log('SKIP chart scenario: requires Windows + Excel')
}

console.log(`conversation turns: report=${reportTurns}, repair=${repairTurns}, analysis=${analysisTurns}, merge=${mergeTurns}, chart=${chartTurns}`)
console.log('ALL CONVERSATION SCENARIOS PASSED')
await ctx.fiber.dispose()
