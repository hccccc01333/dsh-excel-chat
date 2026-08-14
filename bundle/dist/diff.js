import { readFile, writeFile } from 'node:fs/promises';
import { applyPatchesToWorkbook } from './patch.js';
import { readWorkbookCells } from './workbook.js';
export function diffCellMaps(before, after) {
    const ids = new Set([...Object.keys(before), ...Object.keys(after)]);
    const entries = [];
    for (const id of [...ids].sort()) {
        const oldValue = before[id] ?? null;
        const newValue = after[id] ?? null;
        if (oldValue === newValue)
            continue;
        entries.push({
            id,
            kind: oldValue === null ? 'added' : newValue === null ? 'removed' : 'changed',
            oldValue,
            newValue,
        });
    }
    return entries;
}
export function diffToPatches(entries) {
    return entries
        .filter((entry) => entry.kind === 'changed')
        .map((entry) => ({
        id: entry.id,
        kind: 'formula',
        oldValue: entry.oldValue,
        newValue: entry.newValue,
    }));
}
export async function diffWorkbookFiles(beforePath, afterPath) {
    const before = await readWorkbookCells(await readFile(beforePath));
    const after = await readWorkbookCells(await readFile(afterPath));
    return diffCellMaps(before, after);
}
export async function writePatchLog(path, log) {
    await writeFile(path, `${JSON.stringify(log, null, 2)}\n`);
}
export async function readPatchLog(path) {
    const parsed = JSON.parse(await readFile(path, 'utf8'));
    if (parsed.version !== 1 || !Array.isArray(parsed.patches)) {
        throw new Error(`invalid patch log: ${path}`);
    }
    return parsed;
}
export async function applyPatchLog(inputPath, log, outputPath = inputPath) {
    await applyPatchesToWorkbook(inputPath, log.patches, outputPath);
}
export async function rollbackPatchLog(path, log, outPath = path) {
    await applyPatchesToWorkbook(path, log.patches.map((patch) => ({ ...patch, oldValue: patch.newValue, newValue: patch.oldValue })), outPath);
}
