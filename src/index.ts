import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type {} from '@deepseek-ai/dsh-system-prompt'
import Schema from '@deepseek-ai/schemastery'
import { readFile } from 'node:fs/promises'
import { runAgentTask } from './agent.ts'
import { EXCEL_ERROR_VALUES, findErrorCells } from './audit.ts'
import { createLlmRepairAdvisor } from './advisor.ts'
import { autofixWorkbookFile } from './autofix.ts'
import { validateCharts } from './chart-validator.ts'
import {
  createChartWithExcel,
  exportChartsWithExcel,
  exportWorkbookToPdf,
  modifyChartWithExcel,
  type ChartTypeName,
} from './chart-visual.ts'
import { readChartInfos } from './charts.ts'
import { compileFormula } from './compiler.ts'
import { diffWorkbookFiles, readPatchLog, rollbackPatchLog } from './diff.ts'
import { buildDependencyGraph, traceDependencies } from './graph.ts'
import { setLanguage, t } from './i18n.ts'
import { runDoctorChecks } from './doctor.ts'
import { buildWorkbookInsight } from './insight.ts'
import { writeWorkbookHealthReport } from './health-report.ts'
import { explainFormula, readCellContent } from './explain.ts'
import { applyInPlaceEdit, revertInPlaceEdit } from './live-edit.ts'
import { restoreSnapshot } from './safe-write.ts'
import { createLlmPlanner } from './llm-planner.ts'
import { formulaIrSchema } from './ir-schema.ts'
import type { ColumnTable } from './ir.ts'
import { llmTextFromContext } from './llm.ts'
import { excelOperationSchema } from './operation-schema.ts'
import { operateWorkbookFile, type ExcelOperation } from './operations.ts'
import { buildWorkbookMenu } from './menu.ts'
import { registerCommands } from './tools/commands.ts'
import { registerWriteTools } from './tools/write.ts'
import { registerReadTools } from './tools/read.ts'
import { registerDiscoverTools } from './tools/discover.ts'
import { registerChartTools } from './tools/charts.ts'
import { registerAuditTools } from './tools/audit.ts'
import { registerFormulaTools } from './tools/formulas.ts'
import { createPivotTable, type PivotValueSpec } from './pivot.ts'
import { sanitizeOperations } from './plan-schema.ts'
import { buildWorkbookPreview } from './preview.ts'
import { profileWorkbook } from './profile.ts'
import { repairWorkbookFile } from './repair.ts'
import { readWorkbookDetail } from './read.ts'
import { buildWorkbookSemanticProfile } from './semantic.ts'
import { runExcelTask } from './task.ts'
import { detectTableFromCells } from './tables.ts'
import { validate } from './validator.ts'
import { visionTextFromContext } from './vision.ts'
import { readWorkbookCells, readWorkbookSheetNames, validateWorkbookFile } from './workbook.ts'
import { createVisionCritic } from './chart-visual.ts'
import { announce, describeError, guardedContext, registrationSummary, writeStatusReport } from './registration.ts'

export type JsonRecord = Record<string, any>

export const name = 'dsh-excel-chat'
export const inject = ['tools', 'systemPrompt']

export interface ExcelChatConfig {
  /**
   * Language for the text a person reads: health reports, data insights, the
   * capability menu, repair summaries, operation warnings, doctor output.
   *
   * Workbook *data* is deliberately not translated — subtotal labels, generated
   * sheet names and preset names are written into the file and other code keys
   * off them (`patterns.ts` skips summary rows by matching `总计`/`小计`), so
   * translating them would silently change behaviour. The planner prompt is not
   * translated either: it is model input with tuned few-shot examples.
   */
  language?: 'zh' | 'en'
}

/**
 * dsh validates a plugin's config against this schema, so the option shows up
 * in the config surface and a typo is rejected instead of silently ignored.
 * Declared here because the loader reads `Config` from the plugin entry.
 */
export const Config = Schema.object({
  language: Schema.union(['zh', 'en'] as const).default('zh'),
})

export function apply(host: Context, config?: ExcelChatConfig) {
  // Read the config defensively: the loader passes whatever the entry declared,
  // and an unrecognised value must fall back to Chinese rather than throw.
  setLanguage(config?.language === 'en' ? 'en' : 'zh')
  const { ctx, report } = guardedContext(host)
  announce(host, 'info', '[dsh-excel-chat] plugin loaded')
  try {
    registerAll(ctx)
  } catch (error) {
    // Individual registrations are isolated inside `guardedContext`; reaching
    // here means something outside them broke, which would otherwise look
    // exactly like the plugin never loading.
    announce(host, 'error', `[dsh-excel-chat] apply aborted: ${describeError(error)}`)
    throw error
  }
  announce(host, report.failures.length === 0 ? 'info' : 'error', registrationSummary(report))
  writeStatusReport(report)
}

/**
 * Every registration the plugin makes.
 *
 * Kept out of `apply` so that `apply` can report a failure of the registration
 * path itself, separately from a failure of one individual tool.
 */
function registerAll(ctx: Context) {
  registerCommands(ctx)
  ctx.effect(() => ctx.systemPrompt.section({
    name: 'dsh-excel-chat:interaction',
    order: 150,
    text: [
      'Excel 对话交互原则：',
      '- 用户说业务目标而不是操作时（例如“做周报”“帮我整理一下”），不要追问技术细节：调用 excel_menu 给出 2-3 个可选方案让用户挑。',
      '- 用户给出文件但没说要做什么时，先调用 excel_profile 或 excel_menu，主动介绍文件里有什么、列出能做的事，让用户选择。',
      '- 能合理猜出意图时，直接做最可能的版本并展示结果，说明不满意可以用 excel_undo 回滚；删除行列、覆盖数据、删除/保护工作表等破坏性操作必须先确认。',
      '- 运营 / 产品 / 数分岗位用户可以直接套 preset 岗位模板。',
    ].join('\n'),
  }), 'system-prompt:dsh-excel-chat')
  registerWriteTools(ctx)
  registerReadTools(ctx)
  registerDiscoverTools(ctx)
  registerChartTools(ctx)
  registerAuditTools(ctx)
  registerFormulaTools(ctx)
}
