import { t } from '../i18n.js';
import { columnToNumber, normalizeSheet, numberToColumn, parseFormula } from '../formula.js';
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
    shiftSheetRanges(sheet, 'row', start, end, -count);
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
    shiftSheetRanges(sheet, 'col', column, end, -count);
}
/** Replace a formula without losing the cached result ExcelJS keeps alongside it. */
function writeShiftedFormula(cell, formulaWithEquals) {
    const previous = cell.value;
    const result = previous && typeof previous === 'object' ? previous.result : undefined;
    cell.value = (result === undefined
        ? { formula: formulaWithEquals.slice(1) }
        : { formula: formulaWithEquals.slice(1), result });
}
function parseRefSpan(ref) {
    const clean = (ref.split('!').pop() ?? '').replaceAll('$', '');
    const match = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(clean);
    if (!match)
        return null;
    const startCol = columnToNumber(match[1]);
    const startRow = Number(match[2]);
    return {
        startCol,
        startRow,
        endCol: match[3] ? columnToNumber(match[3]) : startCol,
        endRow: match[4] ? Number(match[4]) : startRow,
    };
}
function renderRefSpan(span) {
    const first = `${numberToColumn(span.startCol)}${span.startRow}`;
    if (span.startCol === span.endCol && span.startRow === span.endRow)
        return first;
    return `${first}:${numberToColumn(span.endCol)}${span.endRow}`;
}
/** How a single interval moves when rows/columns are inserted or deleted. */
function shiftInterval(start, end, first, last, delta) {
    if (delta > 0) {
        if (start >= first)
            return [start + delta, end + delta];
        if (end >= first)
            return [start, end + delta];
        return [start, end];
    }
    if (end < first || start > last)
        return start > last ? [start + delta, end + delta] : [start, end];
    const nextStart = start < first ? start : first;
    const nextEnd = end > last ? end + delta : first - 1;
    return nextStart <= nextEnd ? [nextStart, nextEnd] : null;
}
function shiftRefString(ref, axis, first, last, delta) {
    const span = parseRefSpan(ref);
    if (span === null)
        return ref;
    if (axis === 'row') {
        const moved = shiftInterval(span.startRow, span.endRow, first, last, delta);
        if (moved === null)
            return null;
        span.startRow = moved[0];
        span.endRow = moved[1];
    }
    else {
        const moved = shiftInterval(span.startCol, span.endCol, first, last, delta);
        if (moved === null)
            return null;
        span.startCol = moved[0];
        span.endCol = moved[1];
    }
    return renderRefSpan(span);
}
/**
 * ExcelJS's splice only moves cell values and styles. A conditional-format
 * rule, filter, table or merged range left at its old address after a row
 * deletion silently starts judging the wrong data — worse than losing the
 * rule outright. Move every sheet-level range with the edit.
 */
export function shiftSheetRanges(sheet, axis, first, last, delta) {
    const move = (ref) => ref === undefined ? undefined : shiftRefString(ref, axis, first, last, delta);
    // `sheet.model` is a freshly built snapshot on every read. Mutating it does
    // nothing; use the worksheet's live collections instead.
    const cfHolder = sheet;
    if (Array.isArray(cfHolder.conditionalFormattings)) {
        cfHolder.conditionalFormattings = cfHolder.conditionalFormattings
            .map((entry) => ({ ...entry, ref: move(entry.ref) ?? undefined }))
            .filter((entry) => entry.ref !== undefined && entry.ref !== '');
    }
    for (const ref of sheet.model.merges ?? []) {
        const next = shiftRefString(ref, axis, first, last, delta);
        if (next === ref)
            continue;
        sheet.unMergeCells(ref);
        if (next !== null)
            sheet.mergeCells(next);
    }
    // Data validation lives in a worksheet-level address map. Walking cells
    // misses validations on otherwise empty cells, so drive from the map itself.
    const dataValidations = sheet.dataValidations;
    for (const [address, validation] of Object.entries({ ...dataValidations.model })) {
        if (validation === undefined)
            continue;
        const to = shiftRefString(address, axis, first, last, delta);
        if (to === address)
            continue;
        // `remove()` only assigns undefined; the writer still emits the stale
        // address. Delete the key outright so no empty validation survives.
        delete dataValidations.model[address];
        if (to !== null)
            dataValidations.add(to, validation);
    }
    const filter = sheet.autoFilter;
    if (typeof filter === 'string') {
        const shifted = shiftRefString(filter, axis, first, last, delta);
        sheet.autoFilter = (shifted ?? undefined);
    }
    else if (filter) {
        const from = move(filter.from);
        const to = move(filter.to);
        if (from === null || to === null)
            sheet.autoFilter = undefined;
        else {
            if (from !== undefined)
                filter.from = from;
            if (to !== undefined)
                filter.to = to;
        }
    }
    const tables = sheet.tables;
    for (const table of Object.values(tables ?? {})) {
        const next = move(table.ref);
        if (next)
            table.ref = next;
        if (!table.table)
            continue;
        for (const key of ['tableRef', 'autoFilterRef']) {
            const moved = move(table.table[key]);
            if (moved)
                table.table[key] = moved;
        }
    }
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
                    writeShiftedFormula(cell, shifted);
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
                    writeShiftedFormula(cell, shifted);
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
                    writeShiftedFormula(cell, rewritten);
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
                    writeShiftedFormula(cell, rewritten);
            });
        });
    });
}
