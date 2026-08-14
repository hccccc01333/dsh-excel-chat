import { readFile, writeFile } from 'node:fs/promises'
import { applyPatchesToWorkbook, type CellPatch } from './patch.ts'
import { readWorkbookCells } from './workbook.ts'

export interface CellDiffEntry {
  id: string
  kind: 'added' | 'removed' | 'changed'
  oldValue: string | null
  newValue: string | null
}

export function diffCellMaps(before: Record<string, string>, after: Record<string, string>): CellDiffEntry[] {
  const ids = new Set([...Object.keys(before), ...Object.keys(after)])
  const entries: CellDiffEntry[] = []
  for (const id of [...ids].sort()) {
    const oldValue = before[id] ?? null
    const newValue = after[id] ?? null
    if (oldValue === newValue) continue
    entries.push({
      id,
      kind: oldValue === null ? 'added' : newValue === null ? 'removed' : 'changed',
      oldValue,
      newValue,
    })
  }
  return entries
}

export function diffToPatches(entries: CellDiffEntry[]): CellPatch[] {
  return entries
    .filter((entry) => entry.kind === 'changed')
    .map((entry) => ({
      id: entry.id,
      kind: 'formula' as const,
      oldValue: entry.oldValue!,
      newValue: entry.newValue!,
    }))
}

export async function diffWorkbookFiles(beforePath: string, afterPath: string): Promise<CellDiffEntry[]> {
  const before = await readWorkbookCells(await readFile(beforePath))
  const after = await readWorkbookCells(await readFile(afterPath))
  return diffCellMaps(before, after)
}

export interface PatchLog {
  version: 1
  createdAt: string
  sourcePath: string
  patches: CellPatch[]
}

export async function writePatchLog(path: string, log: PatchLog): Promise<void> {
  await writeFile(path, `${JSON.stringify(log, null, 2)}\n`)
}

export async function readPatchLog(path: string): Promise<PatchLog> {
  const parsed = JSON.parse(await readFile(path, 'utf8')) as PatchLog
  if (parsed.version !== 1 || !Array.isArray(parsed.patches)) {
    throw new Error(`invalid patch log: ${path}`)
  }
  return parsed
}

export async function applyPatchLog(inputPath: string, log: PatchLog, outputPath: string = inputPath): Promise<void> {
  await applyPatchesToWorkbook(inputPath, log.patches, outputPath)
}

export async function rollbackPatchLog(path: string, log: PatchLog, outPath: string = path): Promise<void> {
  await applyPatchesToWorkbook(
    path,
    log.patches.map((patch) => ({ ...patch, oldValue: patch.newValue, newValue: patch.oldValue })),
    outPath,
  )
}
