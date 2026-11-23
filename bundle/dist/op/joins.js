import { t } from '../i18n.js';
import { columnToNumber, numberToColumn } from '../formula.js';
import { contentToCellValue } from '../workbook.js';
import { parseRange, shiftFormulaReferences } from './core.js';
export function joinSheets(workbook, operation, warnings, opIndex) {
    if (operation.valueColumns.length !== operation.outputColumns.length) {
        throw new Error(`joinSheets valueColumns (${operation.valueColumns.length}) and outputColumns (${operation.outputColumns.length}) must have the same length`);
    }
    const sourceParsed = parseRange(workbook, operation.source);
    const lookupParsed = parseRange(workbook, operation.lookup);
    const lookupKeyCol = columnToNumber(operation.lookupKey);
    const index = new Map();
    for (let row = lookupParsed.startRow + 1; row <= lookupParsed.endRow; row++) {
        const key = normalizeJoinKey(lookupParsed.sheet.getCell(`${numberToColumn(lookupKeyCol)}${row}`).value);
        if (!key || index.has(key))
            continue;
        index.set(key, operation.valueColumns.map((column) => {
            const col = columnToNumber(column);
            const cell = lookupParsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            return { raw: cell.value, formula: cell.formula ? `=${cell.formula}` : null, row, col };
        }));
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
            const cell = sourceParsed.sheet.getCell(`${numberToColumn(column)}${row}`);
            if (value.formula) {
                // A copied formula must follow its new row, and a cross-sheet lookup
                // must gain the destination sheet as its base for unqualified refs.
                const shifted = shiftFormulaReferences(value.formula, lookupParsed.sheet.name, null, {
                    rowDelta: row - value.row,
                    colDelta: column - value.col,
                });
                cell.value = contentToCellValue(shifted);
            }
            else {
                cell.value = value.raw;
            }
        });
    }
    warnings.push({ op: opIndex, message: t('joinSheets 匹配了 {matched} 行，{missed} 行没有找到对应值', { matched, missed }) });
}
function normalizeJoinKey(value) {
    if (value === null || value === undefined)
        return '';
    return String(typeof value === 'object' && !(value instanceof Date) ? JSON.stringify(value) : value).trim().toLowerCase();
}
