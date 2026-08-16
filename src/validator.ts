import { buildDependencyGraph, type DependencyGraph } from './graph.ts'
import {
  detectEmptyGaps,
  detectErrorValues,
  detectHardcodeBreaks,
  detectPatternAnomalies,
  type ColumnPatternReport,
  type PatternAnomaly,
} from './patterns.ts'

export interface ValidationResult {
  cellCount: number
  formulaCount: number
  dependencyGraph: DependencyGraph
  columns: ColumnPatternReport[]
  anomalies: PatternAnomaly[]
  errors: Array<{ id: string; message: string }>
}

/** Skip plugin-owned internal sheets (e.g. `_dsh_体检报告`) so a health
 * report never flags itself or pollutes user-facing validation. */
function isInternalSheetCell(id: string): boolean {
  const bang = id.lastIndexOf('!')
  if (bang < 0) return false
  return id.slice(0, bang).replace(/^'|'$/g, '').toUpperCase().startsWith('_DSH_')
}

export function validate(cells: Record<string, string>): ValidationResult {
  const owned = Object.fromEntries(
    Object.entries(cells).filter(([id]) => !isInternalSheetCell(id)),
  )
  const formulaEntries: Array<{ id: string; formula: string }> = []
  const errors: Array<{ id: string; message: string }> = []
  for (const [id, content] of Object.entries(owned)) {
    const trimmed = content.trim()
    if (!trimmed.startsWith('=')) continue
    formulaEntries.push({ id, formula: trimmed })
  }

  let dependencyGraph: DependencyGraph
  try {
    dependencyGraph = buildDependencyGraph(formulaEntries)
  } catch (error) {
    errors.push({ id: '*', message: error instanceof Error ? error.message : String(error) })
    dependencyGraph = { edges: [], successors: {}, predecessors: {}, cycles: [] }
  }

  const columns = detectPatternAnomalies(owned)
  const anomalies: PatternAnomaly[] = [
    ...columns.flatMap((column) => column.anomalies),
    ...detectHardcodeBreaks(owned),
    ...detectEmptyGaps(owned),
    ...detectErrorValues(owned),
  ]
  for (const cycle of dependencyGraph.cycles) {
    anomalies.push({
      kind: 'circular-reference',
      cell: cycle[0] ?? '',
      message: `circular reference: ${cycle.join(' -> ')}`,
      expected: 'acyclic',
      actual: cycle.join(' -> '),
      confidence: 1,
    })
  }

  return {
    cellCount: Object.keys(owned).length,
    formulaCount: formulaEntries.length,
    dependencyGraph,
    columns,
    anomalies,
    errors,
  }
}
