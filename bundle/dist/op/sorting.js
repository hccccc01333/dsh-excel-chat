import { columnToNumber, numberToColumn } from '../formula.js';
import { parseRange } from './core.js';
import { normalizeColor } from './formatting.js';
export function sortRange(workbook, range, keys, headerRows) {
    const parsed = parseRange(workbook, range);
    for (const key of keys) {
        if ((key.by === 'fill' || key.by === 'font') && key.color === undefined) {
            throw new Error(`sortRange key with by: "${key.by}" requires color`);
        }
    }
    const keyColumns = keys.map((key) => ({
        column: columnToNumber(key.column),
        direction: key.direction ?? 'asc',
        by: key.by ?? 'value',
        color: key.color === undefined ? undefined : normalizeColor(key.color),
        customList: key.customList,
    }));
    for (const key of keyColumns) {
        if (key.column < parsed.startCol || key.column > parsed.endCol) {
            throw new Error(`sort key column outside range: ${numberToColumn(key.column)}`);
        }
    }
    if (headerRows < 0 || headerRows >= parsed.endRow - parsed.startRow + 1) {
        throw new Error(`invalid headerRows: ${headerRows}`);
    }
    const rows = [];
    for (let row = parsed.startRow + headerRows; row <= parsed.endRow; row++) {
        const cells = {};
        const styles = {};
        const keyValues = [];
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const letter = numberToColumn(col);
            const cell = parsed.sheet.getCell(`${letter}${row}`);
            cells[letter] = cell.value;
            // Carry formatting with the row, the way Excel does — without this a
            // sort by fill colour would move the values out from under the colours.
            styles[letter] = cell.style;
            const key = keyColumns.find((candidate) => candidate.column === col);
            if (key)
                keyValues.push(sortKeyOf(cell, key));
        }
        rows.push({ cells, styles, keys: keyValues });
    }
    rows.sort((a, b) => compareSortRows(a.keys, b.keys, keyColumns.map((key) => key.direction)));
    for (let index = 0; index < rows.length; index++) {
        const targetRow = parsed.startRow + headerRows + index;
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const letter = numberToColumn(col);
            const target = parsed.sheet.getCell(`${letter}${targetRow}`);
            target.value = rows[index].cells[letter] ?? null;
            target.style = (rows[index].styles[letter] ?? {});
        }
    }
}
/**
 * Reduce a cell to whatever its sort key should compare on.
 *
 * Colour keys collapse to a 0/1 group — cells carrying the requested colour sort
 * first, everything else follows — which is what Excel's "move the selected
 * colour to the top" does. Custom lists become the value's position in the list,
 * with unlisted values pushed past the end so they trail the listed ones. Both
 * therefore ride the existing value comparator with no special cases in it.
 */
function sortKeyOf(cell, key) {
    if (key.by === 'fill' || key.by === 'font') {
        const cellColor = key.by === 'fill' ? fillColorOf(cell) : fontColorOf(cell);
        return cellColor !== null && cellColor === key.color ? 0 : 1;
    }
    if (key.customList !== undefined) {
        const index = key.customList.indexOf(String(cell.value ?? ''));
        return index >= 0 ? index : key.customList.length;
    }
    return cell.value;
}
function fillColorOf(cell) {
    const fill = cell.style?.fill;
    if (fill?.type !== 'pattern')
        return null;
    return typeof fill.fgColor?.argb === 'string' ? fill.fgColor.argb.toUpperCase() : null;
}
function fontColorOf(cell) {
    const argb = cell.font?.color?.argb;
    return typeof argb === 'string' ? argb.toUpperCase() : null;
}
function compareSortRows(a, b, directions) {
    for (let index = 0; index < a.length; index++) {
        const comparison = compareSortValue(a[index], b[index]);
        if (comparison !== 0)
            return directions[index] === 'desc' ? -comparison : comparison;
    }
    return 0;
}
function compareSortValue(a, b) {
    if (typeof a === 'number' && typeof b === 'number')
        return a - b;
    if (a instanceof Date && b instanceof Date)
        return a.getTime() - b.getTime();
    const left = a === null || a === undefined ? '' : String(a);
    const right = b === null || b === undefined ? '' : String(b);
    return left < right ? -1 : left > right ? 1 : 0;
}
