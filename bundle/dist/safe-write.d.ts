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
//# sourceMappingURL=safe-write.d.ts.map