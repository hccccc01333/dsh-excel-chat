import { columnToNumber } from './formula.js';
/**
 * Detect a table schema from cell content: the first sheet row with at least
 * two non-numeric, non-formula cells becomes the header row. Text cells map
 * to their column letters. Returns null when no header row is found.
 */
export function detectTableFromCells(cells, sheetName) {
    const rowsBySheet = new Map();
    for (const [id, content] of Object.entries(cells)) {
        const bang = id.lastIndexOf('!');
        const sheet = bang >= 0 ? id.slice(0, bang) : 'Sheet1';
        const cell = bang >= 0 ? id.slice(bang + 1) : id;
        const match = /^([A-Za-z]{1,3})(\d+)$/.exec(cell);
        if (!match)
            continue;
        const text = content.trim();
        if (!text || text.startsWith('=') || /^[+-]?[\d.,%]+$/.test(text))
            continue;
        const row = Number(match[2]);
        if (!rowsBySheet.has(sheet))
            rowsBySheet.set(sheet, new Map());
        const rows = rowsBySheet.get(sheet);
        if (!rows.has(row))
            rows.set(row, []);
        rows.get(row).push({ col: match[1].toUpperCase(), text });
    }
    for (const [sheet, rows] of rowsBySheet) {
        if (sheetName && sheet !== sheetName)
            continue;
        for (const row of [...rows.keys()].sort((a, b) => a - b)) {
            const entries = rows.get(row);
            if (entries.length < 2)
                continue;
            const columns = {};
            for (const entry of entries.sort((a, b) => columnToNumber(a.col) - columnToNumber(b.col))) {
                columns[entry.text] = entry.col;
            }
            return { sheet, columns };
        }
    }
    return null;
}
