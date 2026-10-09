/**
 * Features present in the file that editing it would drop.
 *
 * Returns human-readable names, so a caller can warn precisely instead of saying
 * "something might be lost". An unreadable zip returns an empty list: this is a
 * warning path, and refusing to edit a file because its metadata could not be
 * parsed would be worse than the warning is worth.
 */
export declare function findUnpreservedParts(data: Uint8Array): string[];
export interface SafeWriteResult {
    /** Where the previous contents were copied, when the target already existed. */
    backupPath?: string;
}
/**
 * Write `data` to `path` atomically, backing up whatever was there first.
 *
 * On any failure the temp file is removed and the target is left exactly as it
 * was, so a failed write is a no-op rather than a corruption.
 */
export declare function writeWorkbookSafely(path: string, data: Uint8Array): Promise<SafeWriteResult>;
/** Where the pre-edit copy of `path` lives, when one was taken. */
export declare function snapshotPath(path: string): string;
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
export declare function restoreSnapshot(path: string, outPath?: string): Promise<{
    restored: boolean;
    snapshotPath: string;
}>;
//# sourceMappingURL=safe-write.d.ts.map