import { readFile } from 'node:fs/promises'
import type { CellDiffEntry } from './diff.ts'
import { readWorkbookCells } from './workbook.ts'

export interface WorkbookScore {
  total: number
  matched: number
  mismatched: number
  accuracy: number
  passes: boolean
  mismatches: CellDiffEntry[]
}

export function normalizeCellId(id: string): string {
  const bang = id.lastIndexOf('!')
  if (bang < 0) return `SHEET1!${id.toUpperCase()}`
  const sheet = id.slice(0, bang).replace(/^'|'$/g, '').toUpperCase()
  return `${sheet}!${id.slice(bang + 1).toUpperCase()}`
}

/** Formula comparison tolerance: case, whitespace, and numeric formatting. */
export function cellValueEquals(a: string | null, b: string | null): boolean {
  if (a === b) return true
  if (a === null || b === null) return false
  const left = a.trim()
  const right = b.trim()
  if (left === right) return true
  if (left.startsWith('=') && right.startsWith('=')) {
    return left.toUpperCase().replace(/\s+/g, '') === right.toUpperCase().replace(/\s+/g, '')
  }
  const leftNumber = Number(left)
  const rightNumber = Number(right)
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) {
    return leftNumber === rightNumber
  }
  return false
}

/**
 * Compare a candidate workbook against an oracle workbook cell by cell.
 * Passes when every cell matches (formula case/whitespace and numeric
 * formatting are tolerated). The mismatch entries use the same shape as
 * diffCellMaps(oracle, candidate), so added cells are candidate extras.
 */
export function scoreWorkbookAgainstOracle(
  candidate: Record<string, string>,
  oracle: Record<string, string>,
): WorkbookScore {
  const candidateById = new Map<string, string>()
  for (const [id, content] of Object.entries(candidate)) {
    candidateById.set(normalizeCellId(id), content)
  }
  const oracleById = new Map<string, string>()
  for (const [id, content] of Object.entries(oracle)) {
    oracleById.set(normalizeCellId(id), content)
  }

  const ids = new Set<string>([...candidateById.keys(), ...oracleById.keys()])
  const mismatches: CellDiffEntry[] = []
  let matched = 0
  for (const id of ids) {
    const candidateValue = candidateById.get(id) ?? null
    const oracleValue = oracleById.get(id) ?? null
    if (cellValueEquals(candidateValue, oracleValue)) {
      matched += 1
      continue
    }
    mismatches.push({
      id,
      kind: candidateValue === null ? 'removed' : oracleValue === null ? 'added' : 'changed',
      oldValue: oracleValue,
      newValue: candidateValue,
    })
  }

  const total = ids.size
  return {
    total,
    matched,
    mismatched: total - matched,
    accuracy: total === 0 ? 1 : matched / total,
    passes: matched === total,
    mismatches,
  }
}

export async function scoreWorkbookFiles(candidatePath: string, oraclePath: string): Promise<WorkbookScore> {
  const [candidate, oracle] = await Promise.all([
    readWorkbookCells(await readFile(candidatePath)),
    readWorkbookCells(await readFile(oraclePath)),
  ])
  return scoreWorkbookAgainstOracle(candidate, oracle)
}
