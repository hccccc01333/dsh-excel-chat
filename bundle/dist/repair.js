import { readFile } from 'node:fs/promises';
import { canonicalCellId, columnToNumber, numberToColumn, parseCellId, parseFormula, shiftFormulaRow, } from './formula.js';
import { applyPatchesToWorkbook } from './patch.js';
import { scoreWorkbookAgainstOracle } from './score.js';
import { validate } from './validator.js';
import { readWorkbookCells } from './workbook.js';
export function generateRepairs(cells, result) {
    const anomaliesByCell = new Map();
    for (const column of result.columns) {
        for (const anomaly of column.anomalies) {
            if (anomaly.kind !== 'reference-offset' || !anomaly.slot || !anomaly.expectedOffsets)
                continue;
            if (!anomaliesByCell.has(anomaly.cell))
                anomaliesByCell.set(anomaly.cell, []);
            anomaliesByCell.get(anomaly.cell).push(anomaly);
        }
    }
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
            const refIndex = Number(anomaly.slot.split('.')[0]);
            const ref = parsed.references[refIndex];
            if (!ref)
                continue;
            const base = parseCellId(anomaly.cell);
            const formula = trimmed.slice(1);
            const replacement = ref.end
                ? rebuildRangeText(formula.slice(ref.range.start, ref.range.end), ref, base, anomaliesByCell.get(anomaly.cell) ?? [], refIndex)
                : pointReplacement(ref.start, base, anomaly.expectedOffsets, true);
            if (!replacement)
                continue;
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
    const normalizedCells = new Map();
    for (const [id, content] of Object.entries(cells)) {
        try {
            const parsed = parseCellId(id);
            normalizedCells.set(canonicalCellId(parsed.sheet, parsed.column, parsed.row), content);
        }
        catch {
            // skip malformed ids
        }
    }
    for (const anomaly of result.anomalies) {
        if (anomaly.kind !== 'empty-gap' || repairedCells.has(anomaly.cell))
            continue;
        const base = parseCellId(anomaly.cell);
        const source = findFillSource(normalizedCells, base);
        if (!source)
            continue;
        const newValue = shiftFormulaRow(source.formula, source.rowShift);
        if (!newValue.startsWith('='))
            continue;
        repairs.push({
            id: actualCellId(cells, base, anomaly.cell),
            kind: 'formula',
            oldValue: '',
            newValue,
        });
        repairedCells.add(anomaly.cell);
    }
    return repairs;
}
function findFillSource(normalizedCells, base) {
    for (let distance = 1; distance <= 3; distance++) {
        const aboveId = canonicalCellId(base.sheet, base.column, base.row - distance);
        const above = normalizedCells.get(aboveId)?.trim();
        if (above?.startsWith('='))
            return { formula: above, rowShift: distance };
        const belowId = canonicalCellId(base.sheet, base.column, base.row + distance);
        const below = normalizedCells.get(belowId)?.trim();
        if (below?.startsWith('='))
            return { formula: below, rowShift: -distance };
    }
    return null;
}
/** Reuse the sheet-prefix style of an existing cell on the same sheet. */
function actualCellId(cells, base, fallback) {
    const sample = Object.keys(cells).find((id) => {
        const bang = id.lastIndexOf('!');
        const sheetPart = bang >= 0 ? id.slice(0, bang) : 'Sheet1';
        return sheetPart.replace(/^'|'$/g, '').toUpperCase() === base.sheet;
    });
    if (!sample)
        return fallback;
    const bang = sample.lastIndexOf('!');
    const prefix = bang >= 0 ? `${sample.slice(0, bang)}!` : '';
    return `${prefix}${base.column}${base.row}`;
}
function pointReplacement(point, base, offsets, includeSheetPrefix) {
    if (offsets.colOffset === null || offsets.rowOffset === null)
        return null;
    const columnLetter = numberToColumn(columnToNumber(base.column) + offsets.colOffset);
    const row = base.row + offsets.rowOffset;
    if (!columnLetter || row < 1)
        return null;
    const sheetPrefix = includeSheetPrefix && point.sheet && point.sheet !== base.sheet ? `${point.sheet}!` : '';
    return `${sheetPrefix}${columnLetter}${row}`;
}
/**
 * Rebuild a range reference (e.g. B4:C3) so every deviating endpoint follows
 * the column pattern. An endpoint without an anomaly keeps its original text,
 * including any absolute or sheet prefix.
 */
function rebuildRangeText(rangeText, ref, base, cellAnomalies, refIndex) {
    const colon = rangeText.indexOf(':');
    if (colon < 0 || !ref.end)
        return null;
    const startToken = rangeText.slice(0, colon);
    const endToken = rangeText.slice(colon + 1);
    const startAnomaly = cellAnomalies.find((a) => a.slot === `${refIndex}.start` && a.expectedOffsets);
    const endAnomaly = cellAnomalies.find((a) => a.slot === `${refIndex}.end` && a.expectedOffsets);
    if (!startAnomaly && !endAnomaly)
        return null;
    const newStart = startAnomaly
        ? pointReplacement(ref.start, base, startAnomaly.expectedOffsets, true)
        : startToken;
    const newEnd = endAnomaly
        ? pointReplacement(ref.end, base, endAnomaly.expectedOffsets, endToken.includes('!'))
        : endToken;
    if (!newStart || !newEnd)
        return null;
    return `${newStart}:${newEnd}`;
}
export async function repairWorkbookFile(path, llmAdvisor, cells, oracleCells) {
    const cellMap = cells ?? (await readWorkbookCells(await readFile(path)));
    const before = validate(cellMap);
    const repairs = generateRepairs(cellMap, before);
    const llmRepairs = llmAdvisor ? await llmAdvisor(cellMap, before) : [];
    const covered = new Set(repairs.map((patch) => patch.id));
    const extraLlmRepairs = llmRepairs.filter((patch) => !covered.has(patch.id));
    const allRepairs = [...repairs, ...extraLlmRepairs];
    const repairedPath = path.replace(/\.xlsx$/i, '.repaired.xlsx');
    if (allRepairs.length > 0) {
        await applyPatchesToWorkbook(path, allRepairs, repairedPath);
    }
    const afterCells = allRepairs.length > 0 ? await readWorkbookCells(await readFile(repairedPath)) : cellMap;
    const after = validate(afterCells);
    const oracleScore = oracleCells ? scoreWorkbookAgainstOracle(afterCells, oracleCells) : null;
    return { repairs, llmRepairs: extraLlmRepairs, before, after, repairedPath, oracleScore };
}
