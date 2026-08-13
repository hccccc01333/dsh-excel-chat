import { buildDependencyGraph, type DependencyGraph } from './graph.ts'
import {
  detectEmptyGaps,
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

export function validate(cells: Record<string, string>): ValidationResult {
  const formulaEntries: Array<{ id: string; formula: string }> = []
  const errors: Array<{ id: string; message: string }> = []
  for (const [id, content] of Object.entries(cells)) {
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

  const columns = detectPatternAnomalies(cells)
  const anomalies: PatternAnomaly[] = [
    ...columns.flatMap((column) => column.anomalies),
    ...detectHardcodeBreaks(cells),
    ...detectEmptyGaps(cells),
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
    cellCount: Object.keys(cells).length,
    formulaCount: formulaEntries.length,
    dependencyGraph,
    columns,
    anomalies,
    errors,
  }
}
