export interface LiveEditResult {
    path: string;
    cell: string;
    value: string;
    backupPath: string;
    patchLog: string;
    anomalies: number;
}
/**
 * Apply one cell edit IN PLACE to the local xlsx file (real-time feedback):
 * back up the file to `<path>.bak`, reuse the standard operate pipeline
 * (typed set + formula re-validation + patch audit log), and return the
 * post-edit anomaly count.
 */
export declare function applyInPlaceEdit(path: string, cell: string, value: string | number): Promise<LiveEditResult>;
/**
 * Revert the last in-place edit, restoring the file that was there before it.
 *
 * The snapshot is the whole answer here: an in-place edit always overwrites an
 * existing file, so `writeWorkbookSafely` always left one. Replaying the patch log
 * instead would only put cell values back — the log records value diffs, and a
 * formatting-only edit produces none at all, so undo would report success having
 * changed nothing.
 */
export declare function revertInPlaceEdit(path: string): Promise<{
    restored: boolean;
    message: string;
}>;
//# sourceMappingURL=live-edit.d.ts.map