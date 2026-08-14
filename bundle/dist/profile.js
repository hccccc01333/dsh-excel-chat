import ExcelJS from 'exceljs';
import { readFile } from 'node:fs/promises';
import { numberToColumn } from './formula.js';
import { stripPivotTableParts } from './workbook.js';
/** Default page size suggested by `excel_profile` for chunked `excel_read` calls. */
export const PROFILE_PAGE_SIZE = 100;
/** Rows profiled per sheet; beyond this the stats are sampled and marked truncated. */
const PROFILE_ROW_LIMIT = 5000;
/** Cap on distinct values tracked per column so huge columns stay cheap. */
const UNIQUE_LIMIT = 500;
/** Cap on frequency-map entries per column. */
const TOP_LIMIT = 300;
/**
 * Profile a workbook into a compact structural digest: sheet dimensions,
 * detected header row, per-column dtype/missing/unique/stats/top values, and
 * the range to read first. This is the "structured table encoding" that keeps
 * large sheets readable in one tool call instead of dumping every cell.
 */
export async function profileWorkbook(path, sheet) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(stripPivotTableParts(await readFile(path)));
    const sheets = [];
    for (const worksheet of workbook.worksheets) {
        if (sheet && worksheet.name.toLowerCase() !== sheet.toLowerCase())
            continue;
        sheets.push(profileSheet(worksheet));
    }
    return { path, sheetCount: workbook.worksheets.length, pageSize: PROFILE_PAGE_SIZE, sheets };
}
function profileSheet(worksheet) {
    let maxRow = 0;
    let maxCol = 0;
    let formulaCells = 0;
    const rows = [];
    worksheet.eachRow({ includeEmpty: false }, (row) => {
        const cells = new Map();
        row.eachCell({ includeEmpty: false }, (cell) => {
            // ExcelJS d.ts types col/row as address strings; at runtime they are 1-based numbers.
            const col = cell.col;
            const rowNumber = cell.row;
            cells.set(col, cell);
            maxRow = Math.max(maxRow, rowNumber);
            maxCol = Math.max(maxCol, col);
            if (cell.formula)
                formulaCells++;
        });
        rows.push(cells);
    });
    const headerRow = detectHeaderRow(rows);
    const dataStart = headerRow === null ? 1 : headerRow + 1;
    const dataRows = Math.max(0, maxRow - dataStart + 1);
    const profiledRows = Math.min(dataRows, PROFILE_ROW_LIMIT);
    const truncated = profiledRows < dataRows;
    const columns = [];
    for (let col = 1; col <= maxCol; col++) {
        columns.push(profileColumn(worksheet, rows, col, headerRow, dataStart, profiledRows));
    }
    const readStart = dataRows > 0 ? dataStart : 1;
    const readEndRow = dataRows > 0 ? dataStart + PROFILE_PAGE_SIZE - 1 : maxRow || 1;
    const readHint = `${numberToColumn(1)}${readStart}:${numberToColumn(maxCol || 1)}${Math.min(readEndRow, maxRow || 1)}`;
    return {
        sheet: worksheet.name,
        rowCount: worksheet.rowCount,
        columnCount: worksheet.columnCount,
        usedRange: `${numberToColumn(1)}1:${numberToColumn(maxCol || 1)}${maxRow || 1}`,
        headerRow,
        formulaCells,
        dataRows,
        profiledRows,
        truncated,
        readHint,
        columns,
    };
}
function detectHeaderRow(rows) {
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
        let textCells = 0;
        for (const cell of rows[i].values()) {
            if (cell.formula)
                continue;
            const text = cellValueText(cell);
            if (text && !/^[+-]?[\d.,%]+$/.test(text))
                textCells++;
        }
        if (textCells >= 2)
            return i + 1;
    }
    return null;
}
function profileColumn(worksheet, rows, col, headerRow, dataStart, profiledRows) {
    const column = numberToColumn(col);
    const header = headerRow === null ? null : cellValueText(worksheet.getCell(`${column}${headerRow}`)) || null;
    const seen = new Set();
    const frequencies = new Map();
    const samples = [];
    let uniqueCapped = false;
    let nonEmpty = 0;
    let missing = 0;
    let numericCount = 0;
    let numericSum = 0;
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    let minDate = null;
    let maxDate = null;
    let hasString = false;
    let hasNumber = false;
    let hasDate = false;
    let hasBoolean = false;
    for (let offset = 0; offset < profiledRows; offset++) {
        const rowIndex = dataStart + offset;
        const cell = rows[rowIndex - 1]?.get(col);
        if (!cell) {
            missing++;
            continue;
        }
        const value = cellValue(cell);
        if (value === null || value === undefined || value === '') {
            missing++;
            continue;
        }
        nonEmpty++;
        if (value instanceof Date) {
            hasDate = true;
            if (!minDate || value < minDate)
                minDate = value;
            if (!maxDate || value > maxDate)
                maxDate = value;
        }
        else if (typeof value === 'number') {
            hasNumber = true;
            numericCount++;
            numericSum += value;
            min = Math.min(min, value);
            max = Math.max(max, value);
        }
        else if (typeof value === 'boolean') {
            hasBoolean = true;
        }
        else {
            hasString = true;
        }
        const text = String(value);
        if (seen.size < UNIQUE_LIMIT) {
            seen.add(text);
        }
        else if (!seen.has(text)) {
            uniqueCapped = true;
        }
        if (frequencies.size < TOP_LIMIT) {
            frequencies.set(text, (frequencies.get(text) ?? 0) + 1);
        }
        else if (frequencies.has(text)) {
            frequencies.set(text, (frequencies.get(text) ?? 0) + 1);
        }
        if (samples.length < 3)
            samples.push(text);
    }
    const dtype = !hasString && !hasNumber && !hasDate && !hasBoolean
        ? 'empty'
        : [hasString && 'string', hasNumber && 'number', hasDate && 'date', hasBoolean && 'boolean']
            .filter(Boolean)
            .join(',')
            .includes(',')
            ? 'mixed'
            : hasString ? 'string' : hasNumber ? 'number' : hasDate ? 'date' : 'boolean';
    return {
        column,
        header,
        dtype,
        nonEmpty,
        unique: seen.size,
        uniqueCapped,
        missing,
        ...(hasNumber && Number.isFinite(min) ? { min } : {}),
        ...(hasNumber && Number.isFinite(max) ? { max } : {}),
        ...(numericCount > 0 ? { mean: Math.round((numericSum / numericCount) * 100) / 100 } : {}),
        ...(minDate ? { minDate: minDate.toISOString().slice(0, 10) } : {}),
        ...(maxDate ? { maxDate: maxDate.toISOString().slice(0, 10) } : {}),
        topValues: [...frequencies.entries()]
            .sort((a, b) => b[1] - a[1])
            .slice(0, 3)
            .map(([value, count]) => ({ value, count })),
        samples,
    };
}
function cellValue(cell) {
    if (cell.formula)
        return (cell.result ?? cellValueText(cell)) || null;
    const value = cell.value;
    if (value === null || value === undefined)
        return null;
    if (typeof value === 'object') {
        if (value instanceof Date)
            return value;
        if ('error' in value)
            return String(value.error);
        return cellValueText(cell) || null;
    }
    return value;
}
function cellValueText(cell) {
    const text = cell.text?.trim();
    return text ?? '';
}
