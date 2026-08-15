import { canonicalCellId, columnToNumber, parseCellId, parseFormula } from './formula.js';
function normalizeFormula(parsed, baseSheet, baseColumnNumber, baseRow) {
    const slots = new Map();
    parsed.references.forEach((ref, refIndex) => {
        const points = [['start', ref.start]];
        if (ref.end)
            points.push(['end', ref.end]);
        for (const [role, point] of points) {
            if (!point)
                continue;
            slots.set(`${refIndex}.${role}`, {
                sheet: (point.sheet ?? baseSheet).toUpperCase(),
                colOffset: point.absColumn ? null : columnToNumber(point.column) - baseColumnNumber,
                rowOffset: point.row === null || point.absRow ? null : point.row - baseRow,
            });
        }
    });
    return slots;
}
function normalizedKey(value) {
    return JSON.stringify([value.sheet, value.colOffset, value.rowOffset]);
}
function formatNormalized(value) {
    const parts = [];
    if (value.sheet)
        parts.push(value.sheet);
    parts.push(value.colOffset === null ? 'col:abs' : `col:${formatOffset(value.colOffset)}`);
    parts.push(value.rowOffset === null ? 'row:abs' : `row:${formatOffset(value.rowOffset)}`);
    return parts.join(' ');
}
function formatOffset(offset) {
    if (offset === 0)
        return '0';
    return offset > 0 ? `+${offset}` : `${offset}`;
}
export function detectPatternAnomalies(cells) {
    // Summary rows (subtotal/grand-total labels) legitimately deviate from the
    // column pattern (e.g. a 总计 row sums the summary rows above it). Skip any
    // formula on a row whose label column contains a summary marker.
    const summaryRows = new Set();
    for (const [id, content] of Object.entries(cells)) {
        const trimmed = content.trim();
        if (!/^(总计|小计)$/.test(trimmed) && !trimmed.endsWith('汇总'))
            continue;
        try {
            const cell = parseCellId(id);
            summaryRows.add(`${cell.sheet}:${cell.row}`);
        }
        catch {
            // malformed ids are skipped below
        }
    }
    const groups = new Map();
    for (const [id, content] of Object.entries(cells)) {
        const trimmed = content.trim();
        if (!trimmed.startsWith('='))
            continue;
        // SUBTOTAL rows are aggregate summary rows by design (subtotal/grand total);
        // they follow a different shape than data rows and should not be flagged.
        if (/SUBTOTAL\(/i.test(trimmed))
            continue;
        let cell;
        let parsed;
        try {
            cell = parseCellId(id);
            if (summaryRows.has(`${cell.sheet}:${cell.row}`))
                continue;
            parsed = parseFormula(trimmed);
        }
        catch {
            continue;
        }
        const key = `${cell.sheet}:${cell.column}`;
        if (!groups.has(key))
            groups.set(key, []);
        groups.get(key).push({ id, row: cell.row, parsed });
    }
    const reports = [];
    for (const [key, entries] of groups) {
        if (entries.length < 2)
            continue;
        const separator = key.indexOf(':');
        const sheet = key.slice(0, separator);
        const column = key.slice(separator + 1);
        entries.sort((a, b) => a.row - b.row);
        const baseColumnNumber = columnToNumber(column);
        const normalized = entries.map((entry) => ({
            id: entry.id,
            row: entry.row,
            slots: normalizeFormula(entry.parsed, sheet, baseColumnNumber, entry.row),
        }));
        const slotKeys = new Set();
        for (const entry of normalized) {
            for (const slot of entry.slots.keys())
                slotKeys.add(slot);
        }
        const expected = {};
        const anomalies = [];
        for (const slotKey of slotKeys) {
            const present = normalized.filter((entry) => entry.slots.has(slotKey));
            const counts = new Map();
            for (const entry of present) {
                const value = entry.slots.get(slotKey);
                const bucketKey = normalizedKey(value);
                const bucket = counts.get(bucketKey);
                if (bucket)
                    bucket.count += 1;
                else
                    counts.set(bucketKey, { count: 1, value });
            }
            let majority = { count: -1, value: present[0].slots.get(slotKey) };
            for (const bucket of counts.values()) {
                if (bucket.count > majority.count)
                    majority = bucket;
            }
            const total = present.length;
            expected[slotKey] = formatNormalized(majority.value);
            for (const entry of normalized) {
                const value = entry.slots.get(slotKey);
                if (!value) {
                    // Only cells that miss a majority-shaped slot are anomalies; a slot
                    // carried by fewer than half the cells belongs to a minority shape.
                    if (total >= Math.ceil(normalized.length / 2)) {
                        anomalies.push({
                            kind: 'structure-mismatch',
                            cell: entry.id,
                            message: `missing slot ${slotKey}; column expects ${expected[slotKey]}`,
                            expected: expected[slotKey],
                            actual: null,
                            confidence: total > 0 ? majority.count / total : null,
                        });
                    }
                }
                else if (normalizedKey(value) !== normalizedKey(majority.value)) {
                    anomalies.push({
                        kind: 'reference-offset',
                        cell: entry.id,
                        message: `slot ${slotKey} deviates from column pattern: expected ${expected[slotKey]}, actual ${formatNormalized(value)}`,
                        expected: expected[slotKey],
                        actual: formatNormalized(value),
                        confidence: total > 0 ? majority.count / total : null,
                        slot: slotKey,
                        expectedOffsets: { colOffset: majority.value.colOffset, rowOffset: majority.value.rowOffset },
                        actualOffsets: { colOffset: value.colOffset, rowOffset: value.rowOffset },
                    });
                }
            }
        }
        reports.push({ sheet, column, cellCount: entries.length, expected, anomalies });
    }
    return reports;
}
const NUMERIC_PATTERN = /^[+-]?(\d+(\.\d*)?|\.\d+)(%|e[+-]?\d+)?$/i;
export function detectHardcodeBreaks(cells) {
    const columns = new Map();
    for (const [id, content] of Object.entries(cells)) {
        const trimmed = content.trim();
        if (!trimmed)
            continue;
        let cell;
        try {
            cell = parseCellId(id);
        }
        catch {
            continue;
        }
        const isFormula = trimmed.startsWith('=');
        if (!isFormula && !NUMERIC_PATTERN.test(trimmed))
            continue;
        const key = `${cell.sheet}:${cell.column}`;
        if (!columns.has(key))
            columns.set(key, []);
        columns.get(key).push({ id, row: cell.row, isFormula, content: trimmed });
    }
    const anomalies = [];
    for (const entries of columns.values()) {
        const formulas = entries.filter((entry) => entry.isFormula);
        const values = entries.filter((entry) => !entry.isFormula);
        if (formulas.length < 2 || values.length === 0)
            continue;
        // Subtotal-structured columns interleave summary rows with data rows;
        // numeric cells between SUBTOTAL rows are data, not hardcoded breaks.
        if (formulas.some((entry) => /SUBTOTAL\(/i.test(entry.content)))
            continue;
        const minRow = Math.min(...formulas.map((entry) => entry.row));
        const maxRow = Math.max(...formulas.map((entry) => entry.row));
        for (const value of values) {
            if (value.row >= minRow && value.row <= maxRow) {
                anomalies.push({
                    kind: 'hardcode-break',
                    cell: value.id,
                    message: `numeric value ${value.content} inside formula column (formula rows ${minRow}-${maxRow})`,
                    expected: '=formula',
                    actual: value.content,
                    confidence: null,
                });
            }
        }
    }
    return anomalies;
}
export function detectEmptyGaps(cells) {
    const formulaRows = new Map();
    const knownCells = new Set();
    for (const [id, content] of Object.entries(cells)) {
        const trimmed = content.trim();
        if (!trimmed)
            continue;
        let cell;
        try {
            cell = parseCellId(id);
        }
        catch {
            continue;
        }
        knownCells.add(canonicalCellId(cell.sheet, cell.column, cell.row));
        if (trimmed.startsWith('=')) {
            const key = `${cell.sheet}:${cell.column}`;
            if (!formulaRows.has(key))
                formulaRows.set(key, []);
            formulaRows.get(key).push(cell.row);
        }
    }
    const anomalies = [];
    for (const [key, rows] of formulaRows) {
        const sorted = [...rows].sort((a, b) => a - b);
        const separator = key.indexOf(':');
        const sheet = key.slice(0, separator);
        const column = key.slice(separator + 1);
        for (let i = 1; i < sorted.length; i++) {
            if (sorted[i] - sorted[i - 1] === 2) {
                const gapRow = sorted[i - 1] + 1;
                const gapId = canonicalCellId(sheet, column, gapRow);
                if (!knownCells.has(gapId)) {
                    anomalies.push({
                        kind: 'empty-gap',
                        cell: gapId,
                        message: `empty cell between formula rows ${sorted[i - 1]} and ${sorted[i]}`,
                        expected: `=formula at ${gapId}`,
                        actual: 'empty',
                        confidence: 1,
                    });
                }
            }
        }
    }
    return anomalies;
}
const ERROR_TOKEN = /#(?:REF|DIV\/0|VALUE|NAME\?|N\/A|NULL|NUM)!/g;
/**
 * Detect cells whose content carries an Excel error value such as #REF! or
 * #DIV/0! (both literal error constants and formulas whose cached result is
 * an error token).
 */
export function detectErrorValues(cells) {
    const anomalies = [];
    for (const [id, content] of Object.entries(cells)) {
        const trimmed = content.trim();
        if (!trimmed)
            continue;
        const match = trimmed.match(ERROR_TOKEN);
        if (match) {
            anomalies.push({
                kind: 'error-value',
                cell: id,
                message: `cell contains Excel error ${match[0]}`,
                expected: 'valid value',
                actual: trimmed.slice(0, 200),
                confidence: 1,
            });
        }
    }
    return anomalies;
}
