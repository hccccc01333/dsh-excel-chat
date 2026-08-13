import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { compileFormula } from './compiler.ts'
import { validate } from './validator.ts'
import { validateWorkbookFile } from './workbook.ts'

export const name = 'vera-formula-validator'
export const inject = ['tools']

export function apply(ctx: Context) {
  console.log('[vera-formula-validator] plugin loaded')
  ctx.tools.register(defineTool({
    name: 'excel_compile_formula',
    description: 'Compile a semantic Formula IR into a deterministic Excel formula. IR operations: binary (left/right operands with operator), ratio (numerator/denominator), aggregate (metric, function, filters with value_from cell/column/constant). Table schema: { sheet, columns: { logicalName: columnLetter } }.',
    parameters: {
      ir: {
        type: 'object',
        additionalProperties: true,
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
      return { formula: compileFormula(args.ir, { baseCell: args.baseCell, table: args.table }) }
    },
  }))
  ctx.tools.register(defineTool({
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
        return await validateWorkbookFile(args.path!)
      }
      const normalized: Record<string, string> = {}
      for (const [id, content] of Object.entries(args.cells)) {
        normalized[id] = typeof content === 'string' ? content : String(content)
      }
      return validate(normalized)
    },
  }))
}
