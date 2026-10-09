/**
 * Writing a workbook without risking the one that is already there.
 *
 * Two separate hazards, both of which have already cost data:
 *
 * 1. **A half-written file.** `writeFile` truncates the target and then streams
 *    into it, so a crash, a full disk or a killed process leaves a truncated
 *    `.xlsx` where the user's workbook used to be. Writing to a sibling temp file
 *    and renaming it over the target makes the replacement atomic — the target is
 *    either the old file or the new one, never a fragment. The temp file lives in
 *    the same directory so the rename stays within one filesystem, which is what
 *    makes it atomic at all.
 *
 * 2. **Losing content the writer cannot round-trip.** ExcelJS models cells, styles
 *    and a few objects; it does not model pivot tables, slicers or anything else
 *    it does not know. Those parts vanish from the output. `findUnpreservedParts`
 *    reports them so the caller can say so instead of destroying them silently.
 *
 * A backup is taken whenever the target already exists, so "the write went wrong"
 * is always recoverable rather than final.
 */
import { copyFile, readFile, rename, rm, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { randomUUID } from 'node:crypto'
import { unzipSync } from 'fflate'

/**
 * Content ExcelJS does not model, and therefore cannot write back.
 *
 * Pivot tables and their caches are deliberately absent: `preserve.ts` carries
 * those across a rewrite, so warning about them would be wrong. Everything listed
 * here really is dropped.
 */
const UNPRESERVED = [
  { pattern: /^xl\/slicers\//, feature: 'slicer' },
  { pattern: /^xl\/slicerCaches\//, feature: 'slicer cache' },
  { pattern: /^xl\/timelines\//, feature: 'timeline' },
  { pattern: /^xl\/vbaProject\.bin$/, feature: 'VBA macro project' },
  { pattern: /^xl\/macrosheets\//, feature: 'macro sheet' },
  { pattern: /^xl\/customXml\//, feature: 'custom XML' },
  { pattern: /^xl\/ctrlProps\//, feature: 'form control' },
  { pattern: /^xl\/embeddings\//, feature: 'embedded object' },
  { pattern: /^xl\/diagrams\//, feature: 'SmartArt diagram' },
]

/**
 * Features present in the file that editing it would drop.
 *
 * Returns human-readable names, so a caller can warn precisely instead of saying
 * "something might be lost". An unreadable zip returns an empty list: this is a
 * warning path, and refusing to edit a file because its metadata could not be
 * parsed would be worse than the warning is worth.
 */
export function findUnpreservedParts(data: Uint8Array): string[] {
  let parts: Record<string, Uint8Array>
  try {
    parts = unzipSync(data)
  } catch {
    return []
  }
  const found = new Set<string>()
  for (const name of Object.keys(parts)) {
    for (const { pattern, feature } of UNPRESERVED) {
      if (pattern.test(name)) found.add(feature)
    }
  }
  return [...found]
}

export interface SafeWriteResult {
  /** Where the previous contents were copied, when the target already existed. */
  backupPath?: string
}

/**
 * Write `data` to `path` atomically, backing up whatever was there first.
 *
 * On any failure the temp file is removed and the target is left exactly as it
 * was, so a failed write is a no-op rather than a corruption.
 */
export async function writeWorkbookSafely(path: string, data: Uint8Array): Promise<SafeWriteResult> {
  const temp = join(dirname(path), `.${basename(path)}.writing-${randomUUID()}`)
  await writeFile(temp, data)
  try {
    let backupPath: string | undefined
    if (existsSync(path)) {
      backupPath = `${path}.bak`
      await copyFile(path, backupPath)
    }
    await rename(temp, path)
    return { backupPath }
  } catch (error) {
    await rm(temp, { force: true }).catch(() => {})
    throw error
  }
}

/** Where the pre-edit copy of `path` lives, when one was taken. */
export function snapshotPath(path: string): string {
  return `${path}.bak`
}

/**
 * Restore the pre-edit snapshot over `path`.
 *
 * Undo has to work at the file level, not the cell level. A workbook is rewritten
 * wholesale through ExcelJS on every edit, so an operation changes far more than
 * the cell values a diff can see: styles, merges, row and column structure, sheet
 * order, pivot anchors. Replaying value diffs therefore cannot undo a formatting
 * change — there is nothing in the diff to replay — and it silently reports
 * success. The bytes taken before the write can undo all of it.
 *
 * Returns `restored: false` when no snapshot was taken, which is the caller's cue
 * to fall back to the partial path and say so.
 */
export async function restoreSnapshot(path: string, outPath: string = path): Promise<{ restored: boolean; snapshotPath: string }> {
  const snapshot = snapshotPath(path)
  if (!existsSync(snapshot)) return { restored: false, snapshotPath: snapshot }
  // Going through the safe write keeps the property that a failed restore leaves
  // the current file alone rather than half-replacing it.
  await writeWorkbookSafely(outPath, new Uint8Array(await readFile(snapshot)))
  return { restored: true, snapshotPath: snapshot }
}
