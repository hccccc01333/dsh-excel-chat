import { columnToNumber, normalizeSheet, parseFormula } from '../formula.js';
import { findSheet, shiftFormulaReferences } from './core.js';
/** Delete rows with the same reference-shift semantics as the deleteRows op. */
export function deleteRowsFromSheet(workbook, sheetName, start, count, warnings, opIndex) {
    const sheet = findSheet(workbook, sheetName);
    if (!sheet)
        throw new Error(`sheet not found: ${sheetName}`);
    if (start < 1 || count < 1)
        throw new Error(`invalid deleteRows: row=${start} count=${count}`);
    const end = start + count - 1;
    for (const formulaCell of collectDeletedRangeRefs(workbook, sheetName, start, end)) {
        warnings.push({ op: opIndex, message: `formula ${formulaCell} references a deleted row in ${sheetName}` });
    }
    markDeletedRowRefs(workbook, sheetName, start, end);
    sheet.spliceRows(start, count);
    shiftWorkbookRows(workbook, sheetName, end + 1, -count);
}
/** Delete columns with the same reference-shift semantics as the deleteColumns op. */
export function deleteColumnsFromSheet(workbook, sheetName, column, count, warnings, opIndex) {
    const sheet = findSheet(workbook, sheetName);
    if (!sheet)
        throw new Error(`sheet not found: ${sheetName}`);
    if (column < 1 || count < 1)
        throw new Error(`invalid deleteColumns: column=${column} count=${count}`);
    const end = column + count - 1;
    for (const formulaCell of collectDeletedColumnRefs(workbook, sheetName, column, end)) {
        warnings.push({ op: opIndex, message: `formula ${formulaCell} references a deleted column in ${sheetName}` });
    }
    markDeletedColumnRefs(workbook, sheetName, column, end);
    sheet.spliceColumns(column, count);
    shiftWorkbookColumns(workbook, sheetName, end + 1, -count);
}
export function shiftWorkbookRows(workbook, editedSheet, threshold, rowDelta) {
    const edited = normalizeSheet(editedSheet);
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (!cell.formula)
                    return;
                const formula = `=${cell.formula}`;
                const shifted = shiftFormulaReferences(formula, sheet.name, edited, { rowDelta, rowThreshold: threshold });
                if (shifted !== formula)
                    cell.value = { formula: shifted.slice(1) };
            });
        });
    });
}
export function collectDeletedRangeRefs(workbook, editedSheet, start, end) {
    const edited = normalizeSheet(editedSheet);
    const hits = [];
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (!cell.formula)
                    return;
                const parsed = parseFormula(`=${cell.formula}`);
                for (const ref of parsed.references) {
                    for (const point of [ref.start, ref.end].filter((p) => p !== null)) {
                        const target = normalizeSheet(point.sheet ?? sheet.name);
                        if (target === edited && point.row !== null && !point.absRow && point.row >= start && point.row <= end) {
                            hits.push(`${sheet.name}!${cell.address}`);
                            return;
                        }
                    }
                }
            });
        });
    });
    return hits;
}
export function shiftWorkbookColumns(workbook, editedSheet, threshold, colDelta) {
    const edited = normalizeSheet(editedSheet);
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (!cell.formula)
                    return;
                const formula = `=${cell.formula}`;
                const shifted = shiftFormulaReferences(formula, sheet.name, edited, { colDelta, colThreshold: threshold });
                if (shifted !== formula)
                    cell.value = { formula: shifted.slice(1) };
            });
        });
    });
}
export function collectDeletedColumnRefs(workbook, editedSheet, start, end) {
    const edited = normalizeSheet(editedSheet);
    const hits = [];
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (!cell.formula)
                    return;
                const parsed = parseFormula(`=${cell.formula}`);
                for (const ref of parsed.references) {
                    for (const point of [ref.start, ref.end].filter((p) => p !== null)) {
                        const target = normalizeSheet(point.sheet ?? sheet.name);
                        const columnNumber = columnToNumber(point.column);
                        if (target === edited && !point.absColumn && columnNumber >= start && columnNumber <= end) {
                            hits.push(`${sheet.name}!${cell.address}`);
                            return;
                        }
                    }
                }
            });
        });
    });
    return hits;
}
export function markDeletedRowRefs(workbook, editedSheet, start, end) {
    const edited = normalizeSheet(editedSheet);
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (!cell.formula)
                    return;
                const formula = `=${cell.formula}`;
                const rewritten = shiftFormulaReferences(formula, sheet.name, edited, {
                    rowDeletedStart: start,
                    rowDeletedEnd: end,
                });
                if (rewritten !== formula)
                    cell.value = { formula: rewritten.slice(1) };
            });
        });
    });
}
export function markDeletedColumnRefs(workbook, editedSheet, start, end) {
    const edited = normalizeSheet(editedSheet);
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (!cell.formula)
                    return;
                const formula = `=${cell.formula}`;
                const rewritten = shiftFormulaReferences(formula, sheet.name, edited, {
                    colDeletedStart: start,
                    colDeletedEnd: end,
                });
                if (rewritten !== formula)
                    cell.value = { formula: rewritten.slice(1) };
            });
        });
    });
}
