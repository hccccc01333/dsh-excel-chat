/**
 * Formulas as a subject: plan a task, explain one, compile from IR, validate, trace.
 *
 * The IR is the interesting part — a formula is built from structure rather than
 * from a template string, so it can be checked before it is written. Split out of
 * `index.ts` unchanged.
 */
import type { Context } from '@deepseek-ai/cordis'
import { validate } from '../validator.ts'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { t } from '../i18n.ts'
import { runExcelTask } from '../task.ts'
import type { JsonRecord } from '../index.ts'
import { createLlmPlanner } from '../llm-planner.ts'
import { llmTextFromContext } from '../llm.ts'
import { runAgentTask } from '../agent.ts'
import { explainFormula, readCellContent } from '../explain.ts'
import { formulaIrSchema } from '../ir-schema.ts'
import { compileFormula } from '../compiler.ts'
import type { ColumnTable } from '../ir.ts'
import { readWorkbookCells, validateWorkbookFile } from '../workbook.ts'
import { readFile } from 'node:fs/promises'
import { buildDependencyGraph, traceDependencies } from '../graph.ts'
import { EXCEL_ERROR_VALUES, findErrorCells } from '../audit.ts'

export function registerFormulaTools(ctx: Context): void {
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_task',
  description: 'Execute an Excel workflow in one call. Two modes: (1) steps mode - provide an ordered array of excel_operate-style step operations; after every step the formulas are validated and deterministic repairs are applied, then the next step runs on the verified result. (2) goal mode - provide a natural-language goal and the configured LLM plans the steps, executes them with verification, an LLM verifier checks the goal, and it replans up to maxRounds times until achieved. Use goal mode for vague requests like "make this a monthly report" and steps mode for concrete pipelines like "clean, fill missing, then summarize by region".',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    steps: {
      type: 'array',
      required: true,
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          name: { type: 'string', description: 'Step label, e.g. "clean".' },
          operations: {
            type: 'array',
            required: true,
            items: { type: 'object', additionalProperties: true },
            description: 'excel_operate-style operations for this step (see excel_operate).',
          },
          verify: { type: 'boolean', description: 'Validate and auto-repair formulas after this step (default true).' },
        },
      },
      description: 'Ordered steps; each runs on the previous verified result.',
    },
    goal: {
      type: 'string',
      description: 'Natural-language goal (goal mode). Exactly one of steps or goal must be provided.',
    },
    maxRounds: {
      type: 'number',
      description: 'Maximum plan-execute-verify-replan rounds in goal mode (default 2).',
    },
    provider: {
      type: 'string',
      description: 'LLM provider route for goal mode (default "deepseek").',
    },
    model: {
      type: 'string',
      description: 'LLM model id for goal mode. Required when goal is provided.',
    },
    outPath: {
      type: 'string',
      description: 'Output .xlsx path (default: <path>.task.xlsx).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args) {
    const hasSteps = Array.isArray(args.steps)
    const hasGoal = typeof args.goal === 'string' && args.goal.length > 0
    if (hasSteps === hasGoal) {
      throw new Error('exactly one of steps or goal must be provided')
    }
    const outPath = typeof args.outPath === 'string' && args.outPath ? args.outPath : undefined
    if (hasSteps) {
      return await runExcelTask(
        args.path as string,
        args.steps as Parameters<typeof runExcelTask>[1],
        outPath,
      ) as unknown as JsonRecord
    }
    if (!args.model) {
      throw new Error('model is required when goal is provided')
    }
    const planner = createLlmPlanner(llmTextFromContext(ctx, args.provider ?? 'deepseek', args.model))
    return await runAgentTask(args.path as string, {
      goal: args.goal as string,
      planner,
      maxRounds: typeof args.maxRounds === 'number' && args.maxRounds > 0 ? args.maxRounds : 2,
      outPath,
    }) as unknown as JsonRecord
  },
})), 'tool:excel_task')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_explain_formula',
  description: 'Explain an Excel formula in plain language: parsed functions (SUMIFS / VLOOKUP / IF / date / text / statistics), referenced ranges, cross-sheet references, and arithmetic/comparison. Pass the formula directly, or a path + cell to read it from a workbook. Use when the user asks "这个公式是什么意思".',
  parameters: {
    formula: {
      type: 'string',
      description: 'Formula text, e.g. "=VLOOKUP(A2,Sheet2!$A$1:$B$100,2,FALSE)". Exactly one of formula or path+cell must be provided.',
    },
    path: {
      type: 'string',
      description: 'Absolute path to an .xlsx file (with cell).',
    },
    cell: {
      type: 'string',
      description: 'Cell id to read, e.g. "Sheet1!D4" (with path).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args) {
    const hasFormula = typeof args.formula === 'string' && args.formula.length > 0
    const hasCell = typeof args.path === 'string' && typeof args.cell === 'string'
    if (hasFormula === hasCell) {
      throw new Error('exactly one of formula or path+cell must be provided')
    }
    const formula = hasFormula ? args.formula as string : await readCellContent(args.path as string, args.cell as string)
    if (!formula.trim().startsWith('=')) {
      throw new Error(`cell is not a formula: ${args.cell}`)
    }
    return explainFormula(formula) as unknown as JsonRecord
  },
})), 'tool:excel_explain_formula')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_compile_formula',
  description: 'Compile a semantic Formula IR into a deterministic Excel formula. IR operations: binary (left/right operands with operator), ratio (numerator/denominator), aggregate (metric, function, filters with value_from cell/column/constant). Table schema: { sheet, columns: { logicalName: columnLetter } }.',
  parameters: {
    ir: {
      ...formulaIrSchema,
      required: true,
      description: 'Formula IR object.',
    },
    baseCell: {
      type: 'string',
      required: true,
      description: 'Cell id where the formula will be placed (e.g. "B2" or "Sheet1!B2").',
    },
    table: {
      type: 'object',
      additionalProperties: true,
      required: true,
      description: 'Table schema mapping logical column names to column letters.',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args) {
    return {
      formula: compileFormula(args.ir, {
        baseCell: args.baseCell,
        table: args.table as unknown as ColumnTable,
      }),
    } as unknown as JsonRecord
  },
})), 'tool:excel_compile_formula')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_validate_formulas',
  description: 'Scan workbook cells for silent formula errors: inconsistent reference patterns inside a column, hardcoded values in formula columns, empty fill gaps, and circular references. Pass cells as an object mapping cell id (e.g. "Sheet1!D4" or "D4") to cell content (formulas start with "=").',
  parameters: {
    path: {
      type: 'string',
      description: 'Absolute path to an .xlsx file. Exactly one of path or cells must be provided.',
    },
    cells: {
      type: 'object',
      additionalProperties: true,
      description: 'Map of cell id to cell content. Exactly one of path or cells must be provided.',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args) {
    const hasPath = typeof args.path === 'string' && args.path.length > 0
    const hasCells = args.cells !== undefined
    if (hasPath === hasCells) {
      throw new Error('exactly one of path or cells must be provided')
    }
    if (hasPath) {
      return await validateWorkbookFile(args.path!) as unknown as JsonRecord
    }
    const normalized: Record<string, string> = {}
    for (const [id, content] of Object.entries(args.cells)) {
      normalized[id] = typeof content === 'string' ? content : String(content)
    }
    return validate(normalized) as unknown as JsonRecord
  },
})), 'tool:excel_validate_formulas')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_trace',
  description: 'Trace a cell\'s formula dependencies: precedents (cells it reads) or dependents (cells that read it), breadth-first up to `depth` levels. The data equivalent of Excel\'s Trace Precedents / Trace Dependents — Excel draws arrows, but those are UI state and are never stored in the .xlsx, so the chain is returned as data instead. Each hit carries that cell\'s current value, and circular references found while building the graph are reported too.',
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    cell: {
      type: 'string',
      required: true,
      description: 'Origin cell id, e.g. "Sheet1!D4".',
    },
    direction: {
      type: 'string',
      enum: ['precedents', 'dependents', 'both'],
      description: 'Which way to walk (default both).',
    },
    depth: {
      type: 'number',
      description: 'How many levels to follow (default 1). Use a larger value for the whole chain.',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args) {
    const direction = (args.direction as string | undefined) ?? 'both'
    const depth = args.depth === undefined ? 1 : Number(args.depth)
    const cells = await readWorkbookCells(await readFile(String(args.path)))
    const formulas = Object.entries(cells)
      .map(([id, content]) => ({ id, formula: String(content).trim() }))
      .filter((entry) => entry.formula.startsWith('='))
    const graph = buildDependencyGraph(formulas)
    const cell = String(args.cell)
    // Graph ids are canonical (upper-cased sheet and column); map them back to
    // how the workbook actually spells them so results match what the user sees.
    const displayById = new Map(Object.keys(cells).map((id) => [id.toUpperCase(), id]))
    const display = (id: string): string => displayById.get(id.toUpperCase()) ?? id
    const decorate = (steps: Array<{ cell: string; depth: number }>): JsonRecord[] =>
      steps.map((step) => {
        const id = display(step.cell)
        return { cell: id, depth: step.depth, value: cells[id] ?? null }
      })
    const result: JsonRecord = {
      cell,
      formula: cells[cell] ?? null,
      direction,
      depth,
      cycles: graph.cycles.map((cycle) => cycle.map(display)),
    }
    if (direction === 'precedents' || direction === 'both') {
      const trace = traceDependencies(graph, cell, 'precedents', depth)
      result.precedents = decorate(trace.reached)
      result.precedentsTruncated = trace.truncated
    }
    if (direction === 'dependents' || direction === 'both') {
      const trace = traceDependencies(graph, cell, 'dependents', depth)
      result.dependents = decorate(trace.reached)
      result.dependentsTruncated = trace.truncated
    }
    return result
  },
})), 'tool:excel_trace')
ctx.effect(() => ctx.tools.register(defineTool({
  name: 'excel_find_errors',
  description: `List every cell whose value is an Excel error (${EXCEL_ERROR_VALUES.join(', ')}), together with the formula that produced it and per-code counts. Genuine error values are told apart from text that merely looks like one, so a literal "#N/A" someone typed is not reported.`,
  parameters: {
    path: {
      type: 'string',
      required: true,
      description: 'Absolute path to an .xlsx file.',
    },
    sheet: {
      type: 'string',
      description: 'Limit the scan to one sheet (default: all sheets).',
    },
  },
  output: {
    schema: { type: 'object', additionalProperties: true },
    render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
  },
  async execute(args) {
    const scan = await findErrorCells(
      new Uint8Array(await readFile(String(args.path))),
      args.sheet === undefined ? undefined : String(args.sheet),
    )
    return scan as unknown as JsonRecord
  },
})), 'tool:excel_find_errors')
}
