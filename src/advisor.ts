import { compileFormula } from './compiler.ts'
import type { ColumnTable, FormulaIR } from './ir.ts'
import type { CellPatch } from './patch.ts'
import type { PatternAnomaly } from './patterns.ts'
import type { ValidationResult } from './validator.ts'

export type LlmText = (prompt: string, signal?: AbortSignal) => Promise<string>

export interface LlmRepairReply {
  repairs: Array<{ id: string; baseCell: string; ir: FormulaIR }>
}

export function buildRepairPrompt(
  cells: Record<string, string>,
  anomalies: PatternAnomaly[],
  table: ColumnTable,
): string {
  return [
    'You are the repair planner for a verified Excel agent.',
    'The workbook excerpt (cell id -> content) is:',
    JSON.stringify(cells, null, 2),
    'The validator found these anomalies:',
    JSON.stringify(anomalies, null, 2),
    `The table schema is: ${JSON.stringify(table)}`,
    'Return ONLY JSON matching: {"repairs":[{"id":"<cell id>","baseCell":"<cell id>","ir":<FormulaIR>}]}.',
    'FormulaIR uses operation "binary" (left/right/operator), "ratio" (numerator/denominator),',
    'or "aggregate" (metric/function/filters, each filter with column and value_from).',
  ].join('\n')
}

function stripCodeFence(text: string): string {
  const match = /```(?:json)?\s*([\s\S]*?)```/.exec(text)
  return match ? match[1]!.trim() : text.trim()
}

/**
 * Wrap an LLM text function into a repair advisor: given the workbook excerpt,
 * validation anomalies, and the table schema, ask the model for IR repairs and
 * compile them into concrete CellPatches.
 */
export function createLlmRepairAdvisor(llm: LlmText, table: ColumnTable, signal?: AbortSignal) {
  return async (cells: Record<string, string>, result: ValidationResult): Promise<CellPatch[]> => {
    const anomalies = result.anomalies.filter((anomaly) => anomaly.kind !== 'circular-reference')
    if (anomalies.length === 0) return []
    const text = await llm(buildRepairPrompt(cells, anomalies, table), signal)
    const reply = JSON.parse(stripCodeFence(text)) as LlmRepairReply
    if (!Array.isArray(reply.repairs)) {
      throw new Error('LLM repair reply must contain a repairs array')
    }
    const patches: CellPatch[] = []
    for (const item of reply.repairs) {
      const oldValue = cells[item.id]?.trim()
      if (!oldValue) continue
      patches.push({
        id: item.id,
        kind: 'formula',
        oldValue,
        newValue: compileFormula(item.ir, { baseCell: item.baseCell, table }),
      })
    }
    return patches
  }
}
