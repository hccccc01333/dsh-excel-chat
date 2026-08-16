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
/** Revert the last in-place edit from the patch audit log, restoring the original file. */
export declare function revertInPlaceEdit(path: string): Promise<{
    restored: boolean;
    message: string;
}>;
//# sourceMappingURL=live-edit.d.ts.map