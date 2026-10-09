/**
 * Checking a workbook: chart validation, diffing, formula repair, health report.
 *
 * These are the tools that find problems rather than fix them silently, which is
 * why they are separate from the operations. Split out of `index.ts` unchanged.
 */
import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { t } from '../i18n.ts'
import { readWorkbookCells } from '../workbook.ts'
import { readFile } from 'node:fs/promises'
import { readChartInfos } from '../charts.ts'
import { validateCharts } from '../chart-validator.ts'
import type { JsonRecord } from '../index.ts'
import { diffWorkbookFiles } from '../diff.ts'
import type { ColumnTable } from '../ir.ts'
import { detectTableFromCells } from '../tables.ts'
import { createLlmRepairAdvisor } from '../advisor.ts'
import { llmTextFromContext } from '../llm.ts'
import { repairWorkbookFile } from '../repair.ts'
import { writeWorkbookHealthReport } from '../health-report.ts'
import { autofixWorkbookFile } from '../autofix.ts'

export function registerAuditTools(ctx: Context): void {
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_validate_charts',
  description: 'Validate chart structure inside an .xlsx file: chart type, series references, missing cells, two-dimensional ranges, and unsorted date categories.',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args) {
    const cells = await readWorkbookCells(await readFile(args.path))
    const charts = await readChartInfos(args.path)
    return { charts, reports: validateCharts(charts, cells) } as unknown as JsonRecord
  },
})), 'tool:excel_validate_charts')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_diff_workbook',
  description: 'Compare two .xlsx files cell by cell and return added/removed/changed cells. Useful as a workbook git diff.',
  parameters: {
    beforePath: {
      type: 'string',
      required: true,
      description: 'Absolute path to the original .xlsx file.',
    },
    afterPath: {
      type: 'string',
      required: true,
      description: 'Absolute path to the modified .xlsx file.',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args) {
    return { entries: await diffWorkbookFiles(args.beforePath, args.afterPath) } as unknown as JsonRecord
  },
})), 'tool:excel_diff_workbook')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_repair_formulas',
  description: 'Validate an .xlsx file, generate deterministic repairs for reference-pattern anomalies, optionally ask an LLM to repair remaining anomalies via Formula IR, write a .repaired.xlsx copy, and re-validate.',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    useLlm: {
      type: 'boolean',
      description: 'Ask the configured LLM to repair anomalies the deterministic generator cannot fix.',
    },
    autoTable: {
      type: 'boolean',
      description: 'Detect the header row from cell content when no table schema is provided (uses the first row with two or more text cells).',
    },
    provider: {
      type: 'string',
      description: 'LLM provider route (default "deepseek").',
    },
    model: {
      type: 'string',
      description: 'LLM model id. Required when useLlm is true.',
    },
    table: {
      type: 'object',
      additionalProperties: true,
      description: 'Table schema { sheet, columns } for LLM repair compilation. Required when useLlm is true unless autoTable is enabled.',
    },
    oraclePath: {
      type: 'string',
      description: 'Absolute path to the ground-truth .xlsx file. When provided, the repaired workbook is scored against it and the result includes oracleScore.',
    },
    outPath: {
      type: 'string',
      description: 'Output .xlsx path (default: <path>.repaired.xlsx).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args, exec) {
    let oracleCells: Record<string, string> | undefined
    if (args.oraclePath) {
      oracleCells = await readWorkbookCells(await readFile(args.oraclePath))
    }
    const outputPath = typeof args.outPath === 'string' && args.outPath ? args.outPath : undefined
    if (args.useLlm) {
      if (!args.model) {
        throw new Error('model is required when useLlm is true')
      }
      let table = args.table as unknown as ColumnTable | undefined
      let cells: Record<string, string> | undefined
      if (!table) {
        if (!args.autoTable) {
          throw new Error('table (or autoTable: true) is required when useLlm is true')
        }
        cells = await readWorkbookCells(await readFile(args.path))
        const detected = detectTableFromCells(cells)
        if (!detected) {
          throw new Error('autoTable could not detect a header row; provide table explicitly')
        }
        table = detected
      }
      const advisor = createLlmRepairAdvisor(
        llmTextFromContext(ctx, args.provider ?? 'deepseek', args.model),
        table,
        exec.signal,
      )
      return await repairWorkbookFile(args.path as string, advisor, cells, oracleCells, outputPath) as unknown as JsonRecord
    }
    return await repairWorkbookFile(args.path as string, undefined, undefined, oracleCells, outputPath) as unknown as JsonRecord
  },
})), 'tool:excel_repair_formulas')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_health_report',
  description: 'Write a formula health report INTO the workbook itself: a hidden `_dsh_体检报告` sheet with a health score (100 minus 10 per anomaly), formula/anomaly counts, and one row per anomaly (cell/kind/reason). The file carries its own audit trail. Call after excel_autofix / excel_task or when the user wants the report to travel with the file.',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    outPath: {
      type: 'string',
      description: 'Output .xlsx path (default: write in place).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args) {
    return await writeWorkbookHealthReport(
      args.path as string,
      typeof args.outPath === 'string' && args.outPath ? args.outPath : undefined,
    ) as unknown as JsonRecord
  },
})), 'tool:excel_health_report')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_autofix',
  description: 'One-call self-healing loop for an .xlsx file: validate formulas, apply deterministic repairs for reference-pattern anomalies, optionally ask an LLM to repair the rest via Formula IR, re-validate the repaired copy, and report a plain-language before/after summary. Use after edits or when the user asks to "check and fix" a workbook.',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    useLlm: {
      type: 'boolean',
      description: 'Ask the configured LLM to repair anomalies the deterministic generator cannot fix.',
    },
    autoTable: {
      type: 'boolean',
      description: 'Detect the header row from cell content when no table schema is provided.',
    },
    provider: {
      type: 'string',
      description: 'LLM provider route (default "deepseek").',
    },
    model: {
      type: 'string',
      description: 'LLM model id. Required when useLlm is true.',
    },
    table: {
      type: 'object',
      additionalProperties: true,
      description: 'Table schema { sheet, columns } for LLM repair compilation. Required when useLlm is true unless autoTable is enabled.',
    },
    outPath: {
      type: 'string',
      description: 'Output .xlsx path (default: <path>.repaired.xlsx).',
    },
    healthReport: {
      type: 'boolean',
      description: 'Embed a hidden `_dsh_体检报告` sheet into the repaired file (default true).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args, exec) {
    const outputPath = typeof args.outPath === 'string' && args.outPath ? args.outPath : undefined
    if (!args.useLlm) {
      return await autofixWorkbookFile(args.path as string, {
        outPath: outputPath,
        healthReport: args.healthReport !== false,
      }) as unknown as JsonRecord
    }
    if (!args.model) {
      throw new Error('model is required when useLlm is true')
    }
    let table = args.table as unknown as ColumnTable | undefined
    let cells: Record<string, string> | undefined
    if (!table) {
      if (!args.autoTable) {
        throw new Error('table (or autoTable: true) is required when useLlm is true')
      }
      cells = await readWorkbookCells(await readFile(args.path))
      const detected = detectTableFromCells(cells)
      if (!detected) {
        throw new Error('autoTable could not detect a header row; provide table explicitly')
      }
      table = detected
    }
    const advisor = createLlmRepairAdvisor(
      llmTextFromContext(ctx, args.provider ?? 'deepseek', args.model),
      table,
      exec.signal,
    )
    return await autofixWorkbookFile(args.path as string, {
      advisor,
      outPath: outputPath,
      healthReport: args.healthReport !== false,
    }) as unknown as JsonRecord
  },
})), 'tool:excel_autofix')
}
