import { copyFile, readFile } from 'node:fs/promises';
import { applyOperationsToWorkbook } from './operations.js';
import { diffCellMaps, readPatchLog, rollbackPatchLog, writePatchLog } from './diff.js';
import { validate } from './validator.js';
import { readWorkbookCells } from './workbook.js';
/**
 * Apply one cell edit IN PLACE to the local xlsx file (real-time feedback):
 * back up the file to `<path>.bak`, reuse the standard operate pipeline
 * (typed set + formula re-validation + patch audit log), and return the
 * post-edit anomaly count.
 */
export async function applyInPlaceEdit(path, cell, value) {
    const backupPath = `${path}.bak`;
    await copyFile(path, backupPath);
    const before = await readWorkbookCells(await readFile(path));
    await applyOperationsToWorkbook(path, [{ op: 'set', cells: { [cell]: String(value) } }], path);
    const after = await readWorkbookCells(await readFile(path));
    const patchLogPath = `${path}.patch.json`;
    const log = {
        version: 1,
        createdAt: new Date().toISOString(),
        sourcePath: path,
        patches: diffCellMaps(before, after).map((entry) => ({
            id: entry.id,
            kind: 'formula',
            oldValue: entry.oldValue ?? '',
            newValue: entry.newValue ?? '',
        })),
    };
    await writePatchLog(patchLogPath, log);
    const anomalies = validate(after).anomalies.length;
    return {
        path,
        cell,
        value: String(value),
        backupPath,
        patchLog: patchLogPath,
        anomalies,
    };
}
/** Revert the last in-place edit from the patch audit log, restoring the original file. */
export async function revertInPlaceEdit(path) {
    const patchLogPath = `${path}.patch.json`;
    const log = await readPatchLog(patchLogPath);
    if (log.patches.length === 0) {
        return { restored: false, message: '没有可回滚的编辑记录' };
    }
    await rollbackPatchLog(path, log, path);
    return { restored: true, message: `已回滚 ${log.patches.length} 处修改` };
}
