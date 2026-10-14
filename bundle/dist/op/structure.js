import { t } from '../i18n.js';
import { columnToNumber, normalizeSheet, parseFormula } from '../formula.js';
import { findSheet, requireSheet, shiftFormulaReferences } from './core.js';
/**
 * Reject a start/count pair that is not a positive number.
 *
 * `undefined < 1` is false, so the old inline comparison let a missing `count` through
 * to `sheet.spliceRows(row, undefined)` — which removes nothing and reports no error.
 * The planner path is guarded by the schema, but this layer must not assume it: the
 * whole point of the check is that a malformed call fails loudly here instead of
 * quietly doing nothing.
 */
export function assertPositiveSpan(kind, first, label, count) {
    const start = Number(first);
    const size = Number(count);
    if (!Number.isFinite(start) || start < 1 || !Number.isFinite(size) || size < 1) {
        throw new Error(`invalid ${kind}: ${label}=${String(first)} count=${String(count)}`);
    }
}
/** Delete rows with the same reference-shift semantics as the deleteRows op. */
export function deleteRowsFromSheet(workbook, sheetName, start, count, warnings, opIndex) {
    const sheet = findSheet(workbook, sheetName);
    if (!sheet)
        requireSheet(workbook, sheetName);
    assertPositiveSpan('deleteRows', start, 'row', count);
    const end = start + count - 1;
    for (const formulaCell of collectDeletedRangeRefs(workbook, sheetName, start, end)) {
        warnings.push({ op: opIndex, message: t('公式 {cell} 引用了 {sheet} 中被删除的行', { cell: formulaCell, sheet: sheetName }) });
    }
    markDeletedRowRefs(workbook, sheetName, start, end);
    sheet.spliceRows(start, count);
    shiftWorkbookRows(workbook, sheetName, end + 1, -count);
}
/** Delete columns with the same reference-shift semantics as the deleteColumns op. */
export function deleteColumnsFromSheet(workbook, sheetName, column, count, warnings, opIndex) {
    const sheet = findSheet(workbook, sheetName);
    if (!sheet)
        requireSheet(workbook, sheetName);
    assertPositiveSpan('deleteColumns', column, 'column', count);
    const end = column + count - 1;
    for (const formulaCell of collectDeletedColumnRefs(workbook, sheetName, column, end)) {
        warnings.push({ op: opIndex, message: t('公式 {cell} 引用了 {sheet} 中被删除的列', { cell: formulaCell, sheet: sheetName }) });
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
