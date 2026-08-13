import { readFile } from 'node:fs/promises';
import { columnToNumber, numberToColumn, parseCellId, parseFormula } from './formula.js';
import { applyPatchesToWorkbook } from './patch.js';
import { validate } from './validator.js';
import { readWorkbookCells } from './workbook.js';
export function generateRepairs(cells, result) {
    const repairs = [];
    const repairedCells = new Set();
    for (const column of result.columns) {
        for (const anomaly of column.anomalies) {
            if (anomaly.kind !== 'reference-offset' || !anomaly.slot || !anomaly.expectedOffsets)
                continue;
            if (repairedCells.has(anomaly.cell))
                continue;
            const trimmed = cells[anomaly.cell]?.trim();
            if (!trimmed || !trimmed.startsWith('=') || trimmed.includes('"'))
                continue;
            const parsed = parseFormula(trimmed);
            const [refIndexText, role] = anomaly.slot.split('.');
            if (role !== 'start')
                continue;
            const ref = parsed.references[Number(refIndexText)];
            if (!ref || ref.end)
                continue;
            const base = parseCellId(anomaly.cell);
            const offsets = anomaly.expectedOffsets;
            if (offsets.colOffset === null || offsets.rowOffset === null)
                continue;
            const columnLetter = numberToColumn(columnToNumber(base.column) + offsets.colOffset);
            const row = base.row + offsets.rowOffset;
            if (!columnLetter || row < 1)
                continue;
            const sheetPrefix = ref.start.sheet && ref.start.sheet !== base.sheet ? `${ref.start.sheet}!` : '';
            const replacement = `${sheetPrefix}${columnLetter}${row}`;
            const formula = trimmed.slice(1);
            const rebuilt = `${formula.slice(0, ref.range.start)}${replacement}${formula.slice(ref.range.end)}`;
            repairs.push({
                id: anomaly.cell,
                kind: 'formula',
                oldValue: trimmed,
                newValue: `=${rebuilt}`,
            });
            repairedCells.add(anomaly.cell);
        }
    }
    return repairs;
}
export async function repairWorkbookFile(path, llmAdvisor) {
    const cells = await readWorkbookCells(await readFile(path));
    const before = validate(cells);
    const repairs = generateRepairs(cells, before);
    const llmRepairs = llmAdvisor ? await llmAdvisor(cells, before) : [];
    const covered = new Set(repairs.map((patch) => patch.id));
    const extraLlmRepairs = llmRepairs.filter((patch) => !covered.has(patch.id));
    const allRepairs = [...repairs, ...extraLlmRepairs];
    const repairedPath = path.replace(/\.xlsx$/i, '.repaired.xlsx');
    if (allRepairs.length > 0) {
        await applyPatchesToWorkbook(path, allRepairs, repairedPath);
    }
    const afterCells = allRepairs.length > 0 ? await readWorkbookCells(await readFile(repairedPath)) : cells;
    const after = validate(afterCells);
    return { repairs, llmRepairs: extraLlmRepairs, before, after, repairedPath };
}
