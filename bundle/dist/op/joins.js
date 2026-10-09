import { columnToNumber, numberToColumn } from '../formula.js';
import { contentToCellValue } from '../workbook.js';
import { cellContentOf, parseRange } from './core.js';
export function joinSheets(workbook, operation, warnings, opIndex) {
    if (operation.valueColumns.length !== operation.outputColumns.length) {
        throw new Error(`joinSheets valueColumns (${operation.valueColumns.length}) and outputColumns (${operation.outputColumns.length}) must have the same length`);
    }
    const sourceParsed = parseRange(workbook, operation.source);
    const lookupParsed = parseRange(workbook, operation.lookup);
    const lookupKeyCol = columnToNumber(operation.lookupKey);
    // First match wins, mirroring VLOOKUP's approximate=false behaviour.
    const index = new Map();
    for (let row = lookupParsed.startRow + 1; row <= lookupParsed.endRow; row++) {
        const key = normalizeJoinKey(lookupParsed.sheet.getCell(`${numberToColumn(lookupKeyCol)}${row}`).value);
        if (!key || index.has(key))
            continue;
        index.set(key, operation.valueColumns.map((column) => cellContentOf(lookupParsed.sheet.getCell(`${numberToColumn(columnToNumber(column))}${row}`))));
    }
    const sourceKeyCol = columnToNumber(operation.sourceKey);
    let matched = 0;
    let missed = 0;
    for (let row = sourceParsed.startRow + 1; row <= sourceParsed.endRow; row++) {
        const key = normalizeJoinKey(sourceParsed.sheet.getCell(`${numberToColumn(sourceKeyCol)}${row}`).value);
        const values = key ? index.get(key) : undefined;
        if (!values) {
            missed++;
            if (operation.missValue !== undefined) {
                operation.outputColumns.forEach((column, i) => {
                    sourceParsed.sheet.getCell(`${numberToColumn(columnToNumber(column))}${row}`).value =
                        typeof operation.missValue === 'number' ? operation.missValue : String(operation.missValue ?? '');
                });
            }
            continue;
        }
        matched++;
        values.forEach((value, i) => {
            const column = columnToNumber(operation.outputColumns[i]);
            sourceParsed.sheet.getCell(`${numberToColumn(column)}${row}`).value =
                value.startsWith('=') ? { formula: value.slice(1) } : contentToCellValue(value);
        });
    }
    warnings.push({ op: opIndex, message: `joinSheets matched ${matched} row(s), ${missed} without a lookup hit` });
}
function normalizeJoinKey(value) {
    if (value === null || value === undefined)
        return '';
    return String(typeof value === 'object' && !(value instanceof Date) ? JSON.stringify(value) : value).trim().toLowerCase();
}
