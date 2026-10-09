import { copyFile, readFile } from 'node:fs/promises';
import { applyOperationsToWorkbook } from './operations.js';
import { diffCellMaps, readPatchLog, rollbackPatchLog, writePatchLog } from './diff.js';
import { t } from './i18n.js';
import { restoreSnapshot } from './safe-write.js';
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
/**
 * Revert the last in-place edit, restoring the file that was there before it.
 *
 * The snapshot is the whole answer here: an in-place edit always overwrites an
 * existing file, so `writeWorkbookSafely` always left one. Replaying the patch log
 * instead would only put cell values back — the log records value diffs, and a
 * formatting-only edit produces none at all, so undo would report success having
 * changed nothing.
 */
export async function revertInPlaceEdit(path) {
    const { restored } = await restoreSnapshot(path);
    // The log may not exist at all; a missing audit trail is not a reason to fail.
    const log = await readPatchLog(`${path}.patch.json`)
        .catch(() => ({ version: 1, createdAt: '', sourcePath: path, patches: [] }));
    if (restored) {
        return { restored: true, message: t('已按编辑前的完整快照恢复（{count} 处改动，含格式与结构）', { count: log.patches.length }) };
    }
    if (log.patches.length === 0) {
        return { restored: false, message: t('没有可回滚的编辑记录') };
    }
    await rollbackPatchLog(path, log, path);
    return { restored: true, message: t('没有编辑前的快照，只能按审计日志还原 {count} 处单元格内容（格式与结构不在日志里）', { count: log.patches.length }) };
}
