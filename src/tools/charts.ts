/**
 * Charts: create, modify, export to PNG or PDF, and check them visually.
 *
 * Chart work needs Excel on Windows for the parts ExcelJS cannot do, which is why
 * these report capability rather than assuming it. Split out of `index.ts` unchanged.
 */
import type { Context } from '@deepseek-ai/cordis'
import {
  createChartWithExcel,
  exportChartsWithExcel,
  exportWorkbookToPdf,
  modifyChartWithExcel,
  type ChartTypeName,
} from '../chart-visual.ts'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { t } from '../i18n.ts'
import { createVisionCritic } from '../chart-visual.ts'
import { visionTextFromContext } from '../vision.ts'
import type { JsonRecord } from '../index.ts'

export function registerChartTools(ctx: Context): void {
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_validate_charts_visual',
  description: 'Export all charts from an .xlsx file to PNG using local Excel, then ask the configured vision-capable LLM to check visual quality (title, legend, labels, axes, crowding, trend readability).',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    model: {
      type: 'string',
      required: true,
      description: 'Vision-capable LLM model id.',
    },
    provider: {
      type: 'string',
      description: 'LLM provider route (default "deepseek").',
    },
    outDir: {
      type: 'string',
      description: 'Output directory for PNG files (default: <path>.charts).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args, exec) {
    const outDir = args.outDir ?? `${args.path.replace(/\.xlsx$/i, '')}.charts`
    const images = await exportChartsWithExcel(args.path, outDir, exec.signal)
    const critic = createVisionCritic(visionTextFromContext(ctx, args.provider ?? 'deepseek', args.model))
    const reports = []
    for (const image of images) {
      reports.push(await critic(image, exec.signal))
    }
    return { images, reports } as unknown as JsonRecord
  },
})), 'tool:excel_validate_charts_visual')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_create_chart',
  description: 'Create a chart in an .xlsx copy using local Excel (Windows only): choose data range, chart type (column/line/pie/bar/area), title, and chart name.',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    range: {
      type: 'string',
      required: true,
      description: 'Data range including headers, e.g. "Sheet1!A1:B4".',
    },
    sheet: {
      type: 'string',
      description: 'Sheet name containing the range (default first sheet).',
    },
    type: {
      type: 'string',
      enum: ['column', 'line', 'pie', 'bar', 'area'],
      description: 'Chart type (default column).',
    },
    title: {
      type: 'string',
      description: 'Chart title.',
    },
    name: {
      type: 'string',
      description: 'Chart name (default "Chart 1").',
    },
    outPath: {
      type: 'string',
      description: 'Output .xlsx path (default: <path>.chart.xlsx).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args, exec) {
    if (process.platform !== 'win32') {
      throw new Error('chart creation requires Windows with Microsoft Excel installed')
    }
    const outPath = (typeof args.outPath === 'string' && args.outPath ? args.outPath : (args.path as string).replace(/\.xlsx$/i, '.chart.xlsx'))
    await createChartWithExcel(args.path as string, {
      sheet: typeof args.sheet === 'string' ? args.sheet : undefined,
      range: args.range as string,
      type: args.type as ChartTypeName | undefined,
      title: typeof args.title === 'string' ? args.title : undefined,
      name: typeof args.name === 'string' ? args.name : undefined,
    }, outPath, exec.signal)
    return { outputPath: outPath } as unknown as JsonRecord
  },
})), 'tool:excel_create_chart')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_modify_chart',
  description: 'Modify chart parameters in an .xlsx copy using local Excel (Windows only): chart type, title, legend, and axis titles.',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    chartName: {
      type: 'string',
      required: true,
      description: 'Name of the chart to modify (e.g. "Chart 1").',
    },
    type: {
      type: 'string',
      enum: ['column', 'line', 'pie', 'bar', 'area'],
      description: 'New chart type.',
    },
    title: {
      type: 'string',
      description: 'New chart title.',
    },
    hasLegend: {
      type: 'boolean',
      description: 'Show or hide the legend.',
    },
    axisTitleX: {
      type: 'string',
      description: 'Category axis title.',
    },
    axisTitleY: {
      type: 'string',
      description: 'Value axis title.',
    },
    outPath: {
      type: 'string',
      description: 'Output .xlsx path (default: <path>.chart.xlsx).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args, exec) {
    if (process.platform !== 'win32') {
      throw new Error('chart modification requires Windows with Microsoft Excel installed')
    }
    const outPath = (typeof args.outPath === 'string' && args.outPath ? args.outPath : (args.path as string).replace(/\.xlsx$/i, '.chart.xlsx'))
    await modifyChartWithExcel(args.path as string, args.chartName as string, {
      type: args.type as ChartTypeName | undefined,
      title: typeof args.title === 'string' ? args.title : undefined,
      hasLegend: typeof args.hasLegend === 'boolean' ? args.hasLegend : undefined,
      axisTitleX: typeof args.axisTitleX === 'string' ? args.axisTitleX : undefined,
      axisTitleY: typeof args.axisTitleY === 'string' ? args.axisTitleY : undefined,
    }, outPath, exec.signal)
    return { outputPath: outPath } as unknown as JsonRecord
  },
})), 'tool:excel_modify_chart')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_export_charts',
  description: 'Export all charts from an .xlsx file to PNG images using local Microsoft Excel (Windows only).',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    outDir: {
      type: 'string',
      description: 'Output directory for PNG files (default: <path>.charts).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args, exec) {
    const outDir = args.outDir ?? `${args.path.replace(/\.xlsx$/i, '')}.charts`
    return { images: await exportChartsWithExcel(args.path, outDir, exec.signal) } as unknown as JsonRecord
  },
})), 'tool:excel_export_charts')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_export_pdf',
  description: 'Export an .xlsx file (whole workbook or one sheet) to a PDF using local Microsoft Excel COM. Windows only; the source file is opened read-only and left untouched.',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    outPath: {
      type: 'string',
      description: 'Output PDF path (default: <path>.pdf).',
    },
    sheet: {
      type: 'string',
      description: 'Export only this sheet (default all sheets).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args, exec) {
    const outPath = typeof args.outPath === 'string' && args.outPath
      ? args.outPath
      : (args.path as string).replace(/\.xlsx$/i, '.pdf')
    await exportWorkbookToPdf(args.path as string, outPath, typeof args.sheet === 'string' && args.sheet ? args.sheet : undefined, exec.signal)
    return { pdfPath: outPath } as unknown as JsonRecord
  },
})), 'tool:excel_export_pdf')
}
