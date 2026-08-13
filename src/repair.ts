import { readFile } from 'node:fs/promises'
import {
  columnToNumber,
  numberToColumn,
  parseCellId,
  parseFormula,
  type ParsedCellId,
  type ParsedRef,
  type RefPoint,
} from './formula.ts'
import { applyPatchesToWorkbook, type CellPatch } from './patch.ts'
import type { PatternAnomaly } from './patterns.ts'
import { validate, type ValidationResult } from './validator.ts'
import { readWorkbookCells } from './workbook.ts'

export type RepairAdvisor = (cells: Record<string, string>, result: ValidationResult) => Promise<CellPatch[]>

export function generateRepairs(cells: Record<string, string>, result: ValidationResult): CellPatch[] {
  const anomaliesByCell = new Map<string, PatternAnomaly[]>()
  for (const column of result.columns) {
    for (const anomaly of column.anomalies) {
      if (anomaly.kind !== 'reference-offset' || !anomaly.slot || !anomaly.expectedOffsets) continue
      if (!anomaliesByCell.has(anomaly.cell)) anomaliesByCell.set(anomaly.cell, [])
      anomaliesByCell.get(anomaly.cell)!.push(anomaly)
    }
  }

  const repairs: CellPatch[] = []
  const repairedCells = new Set<string>()
  for (const column of result.columns) {
    for (const anomaly of column.anomalies) {
      if (anomaly.kind !== 'reference-offset' || !anomaly.slot || !anomaly.expectedOffsets) continue
      if (repairedCells.has(anomaly.cell)) continue
      const trimmed = cells[anomaly.cell]?.trim()
      if (!trimmed || !trimmed.startsWith('=') || trimmed.includes('"')) continue
      const parsed = parseFormula(trimmed)
      const refIndex = Number(anomaly.slot.split('.')[0])
      const ref = parsed.references[refIndex]
      if (!ref) continue
      const base = parseCellId(anomaly.cell)
      const formula = trimmed.slice(1)
      const replacement = ref.end
        ? rebuildRangeText(
            formula.slice(ref.range.start, ref.range.end),
            ref,
            base,
            anomaliesByCell.get(anomaly.cell) ?? [],
            refIndex,
          )
        : pointReplacement(ref.start, base, anomaly.expectedOffsets, true)
      if (!replacement) continue
      const rebuilt = `${formula.slice(0, ref.range.start)}${replacement}${formula.slice(ref.range.end)}`
      repairs.push({
        id: anomaly.cell,
        kind: 'formula',
        oldValue: trimmed,
        newValue: `=${rebuilt}`,
      })
      repairedCells.add(anomaly.cell)
    }
  }
  return repairs
}

function pointReplacement(
  point: RefPoint,
  base: ParsedCellId,
  offsets: { colOffset: number | null; rowOffset: number | null },
  includeSheetPrefix: boolean,
): string | null {
  if (offsets.colOffset === null || offsets.rowOffset === null) return null
  const columnLetter = numberToColumn(columnToNumber(base.column) + offsets.colOffset)
  const row = base.row + offsets.rowOffset
  if (!columnLetter || row < 1) return null
  const sheetPrefix = includeSheetPrefix && point.sheet && point.sheet !== base.sheet ? `${point.sheet}!` : ''
  return `${sheetPrefix}${columnLetter}${row}`
}

/**
 * Rebuild a range reference (e.g. B4:C3) so every deviating endpoint follows
 * the column pattern. An endpoint without an anomaly keeps its original text,
 * including any absolute or sheet prefix.
 */
function rebuildRangeText(
  rangeText: string,
  ref: ParsedRef,
  base: ParsedCellId,
  cellAnomalies: PatternAnomaly[],
  refIndex: number,
): string | null {
  const colon = rangeText.indexOf(':')
  if (colon < 0 || !ref.end) return null
  const startToken = rangeText.slice(0, colon)
  const endToken = rangeText.slice(colon + 1)
  const startAnomaly = cellAnomalies.find((a) => a.slot === `${refIndex}.start` && a.expectedOffsets)
  const endAnomaly = cellAnomalies.find((a) => a.slot === `${refIndex}.end` && a.expectedOffsets)
  if (!startAnomaly && !endAnomaly) return null
  const newStart = startAnomaly
    ? pointReplacement(ref.start, base, startAnomaly.expectedOffsets!, true)
    : startToken
  const newEnd = endAnomaly
    ? pointReplacement(ref.end, base, endAnomaly.expectedOffsets!, endToken.includes('!'))
    : endToken
  if (!newStart || !newEnd) return null
  return `${newStart}:${newEnd}`
}

export interface RepairResult {
  repairs: CellPatch[]
  llmRepairs: CellPatch[]
  before: ValidationResult
  after: ValidationResult
  repairedPath: string
}

export async function repairWorkbookFile(
  path: string,
  llmAdvisor?: RepairAdvisor,
  cells?: Record<string, string>,
): Promise<RepairResult> {
  const cellMap = cells ?? (await readWorkbookCells(await readFile(path)))
  const before = validate(cellMap)
  const repairs = generateRepairs(cellMap, before)
  const llmRepairs = llmAdvisor ? await llmAdvisor(cellMap, before) : []
  const covered = new Set(repairs.map((patch) => patch.id))
  const extraLlmRepairs = llmRepairs.filter((patch) => !covered.has(patch.id))
  const allRepairs = [...repairs, ...extraLlmRepairs]
  const repairedPath = path.replace(/\.xlsx$/i, '.repaired.xlsx')
  if (allRepairs.length > 0) {
    await applyPatchesToWorkbook(path, allRepairs, repairedPath)
  }
  const afterCells = allRepairs.length > 0 ? await readWorkbookCells(await readFile(repairedPath)) : cellMap
  const after = validate(afterCells)
  return { repairs, llmRepairs: extraLlmRepairs, before, after, repairedPath }
}
