import ExcelJS from 'exceljs';
import { columnToNumber, normalizeSheet, numberToColumn, parseCellId, parseFormula, } from './formula.js';
import { validate } from './validator.js';
import { readWorkbookCells } from './workbook.js';
import { readFile } from 'node:fs/promises';
const RANGE_LINE = /^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/;
export function findSheet(workbook, name) {
    const normalized = normalizeSheet(name);
    return workbook.worksheets.find((sheet) => normalizeSheet(sheet.name) === normalized);
}
function resolveCell(workbook, id) {
    const parsed = parseCellId(id);
    const sheet = findSheet(workbook, parsed.sheet);
    if (!sheet)
        throw new Error(`sheet not found: ${parsed.sheet}`);
    return sheet.getCell(`${parsed.column}${parsed.row}`);
}
function writeContent(cell, content) {
    const trimmed = content.trim();
    cell.value = toCellValue(trimmed);
}
/**
 * Convert user-provided text into an Excel value: formulas stay formulas,
 * plain numbers/dates/booleans keep their type, everything else is text.
 * Workplace spreadsheets break when "100" is written as text, so numeric
 * strings are typed before they reach ExcelJS.
 */
function toCellValue(content) {
    if (content.startsWith('='))
        return { formula: content.slice(1) };
    if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(content))
        return Number(content);
    if (/^true$/i.test(content))
        return true;
    if (/^false$/i.test(content))
        return false;
    const date = /^(\d{4})-(\d{2})-(\d{2})([T ](\d{2}):(\d{2})(:(\d{2}))?)?$/.exec(content);
    if (date) {
        const year = Number(date[1]);
        const month = Number(date[2]);
        const day = Number(date[3]);
        const hour = date[5] ? Number(date[5]) : 0;
        const minute = date[6] ? Number(date[6]) : 0;
        const second = date[8] ? Number(date[8]) : 0;
        return new Date(year, month - 1, day, hour, minute, second);
    }
    return content;
}
function parseRange(workbook, range) {
    const bang = range.lastIndexOf('!');
    if (bang < 0)
        throw new Error(`range requires a sheet: ${range}`);
    const rawSheet = range.slice(0, bang);
    const body = range.slice(bang + 1);
    const match = RANGE_LINE.exec(body);
    if (!match)
        throw new Error(`invalid range: ${range}`);
    const sheet = findSheet(workbook, rawSheet);
    if (!sheet)
        throw new Error(`sheet not found: ${rawSheet}`);
    return {
        sheet,
        startCol: columnToNumber(match[1]),
        startRow: Number(match[2]),
        endCol: columnToNumber(match[3]),
        endRow: Number(match[4]),
    };
}
/**
 * Shift selected reference points of a formula. rowDelta/colDelta apply to
 * relative rows/columns; rowThreshold/colThreshold gate the shift so row edits
 * only move references at or below the insertion/deletion point. When
 * editedSheet is set, only references pointing into that sheet are shifted.
 */
export function shiftFormulaReferences(formula, baseSheet, editedSheet, options = {}) {
    const { rowDelta, colDelta, rowThreshold, colThreshold, rowDeletedStart, rowDeletedEnd, colDeletedStart, colDeletedEnd } = options;
    if ((rowDelta ?? 0) === 0 &&
        (colDelta ?? 0) === 0 &&
        rowDeletedStart === undefined &&
        colDeletedStart === undefined)
        return formula;
    const hasEquals = formula.trimStart().startsWith('=');
    const raw = hasEquals ? formula.trimStart().slice(1) : formula;
    const parsed = parseFormula(`=${raw}`);
    const edits = [];
    for (const ref of parsed.references) {
        const text = raw.slice(ref.range.start, ref.range.end);
        const colon = text.indexOf(':');
        const startToken = colon >= 0 ? text.slice(0, colon) : text;
        const endToken = colon >= 0 ? text.slice(colon + 1) : null;
        const newStart = shiftPointToken(startToken, ref.start, baseSheet, editedSheet, { rowDelta, colDelta, rowThreshold, colThreshold, rowDeletedStart, rowDeletedEnd, colDeletedStart, colDeletedEnd });
        if (endToken === null) {
            edits.push({ start: ref.range.start, end: ref.range.end, text: newStart });
        }
        else {
            const newEnd = shiftPointToken(endToken, ref.end, baseSheet, editedSheet, { rowDelta, colDelta, rowThreshold, colThreshold, rowDeletedStart, rowDeletedEnd, colDeletedStart, colDeletedEnd });
            edits.push({ start: ref.range.start, end: ref.range.end, text: `${newStart}:${newEnd}` });
        }
    }
    let result = raw;
    for (const edit of [...edits].sort((a, b) => b.start - a.start)) {
        result = `${result.slice(0, edit.start)}${edit.text}${result.slice(edit.end)}`;
    }
    return hasEquals ? `=${result}` : result;
}
function shiftPointToken(token, point, baseSheet, editedSheet, options) {
    const effectiveSheet = normalizeSheet(point.sheet ?? baseSheet);
    if (editedSheet && effectiveSheet !== normalizeSheet(editedSheet))
        return token;
    const { rowDelta, colDelta, rowThreshold, colThreshold, rowDeletedStart, rowDeletedEnd, colDeletedStart, colDeletedEnd, } = options;
    const colMatch = /^(.*?)(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(token);
    const wholeColMatch = /^(.*?)(\$?)([A-Za-z]{1,3})$/.exec(token);
    const hasRow = colMatch !== null;
    const prefix = hasRow ? colMatch[1] : wholeColMatch?.[1] ?? token;
    const absCol = hasRow ? colMatch[2] === '$' : wholeColMatch?.[2] === '$';
    const col = hasRow ? colMatch[3] : wholeColMatch?.[3] ?? null;
    const absRow = hasRow ? colMatch[4] === '$' : false;
    const row = hasRow ? Number(colMatch[5]) : null;
    const currentColNumber = col ? columnToNumber(col) : null;
    if ((rowDeletedStart !== undefined && row !== null && !absRow && row >= rowDeletedStart && row <= (rowDeletedEnd ?? rowDeletedStart)) ||
        (colDeletedStart !== undefined && currentColNumber !== null && !absCol && currentColNumber >= colDeletedStart && currentColNumber <= (colDeletedEnd ?? colDeletedStart))) {
        return '#REF!';
    }
    let newCol = col;
    if (col && !absCol && colDelta) {
        const current = columnToNumber(col);
        if ((colThreshold === undefined || current >= colThreshold) && current + colDelta >= 1) {
            newCol = numberToColumn(current + colDelta);
        }
    }
    let newRow = row;
    if (row !== null && !absRow && rowDelta) {
        if ((rowThreshold === undefined || row >= rowThreshold) && row + rowDelta >= 1) {
            newRow = row + rowDelta;
        }
    }
    if (newCol === col && newRow === row)
        return token;
    const colPart = `${absCol ? '$' : ''}${newCol ?? ''}`;
    const rowPart = newRow === null ? '' : `${absRow ? '$' : ''}${newRow}`;
    return `${prefix}${colPart}${rowPart}`;
}
function cellContentOf(cell) {
    if (cell.formula)
        return `=${cell.formula}`;
    const value = cell.value;
    if (value === null || value === undefined)
        return '';
    return typeof value === 'object' ? JSON.stringify(value) : String(value);
}
export async function applyOperationsToWorkbook(inputPath, operations, outputPath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputPath);
    const warnings = [];
    operations.forEach((operation, index) => {
        switch (operation.op) {
            case 'set': {
                for (const [id, content] of Object.entries(operation.cells)) {
                    writeContent(resolveCell(workbook, id), content);
                }
                break;
            }
            case 'fill': {
                applyFill(workbook, operation.source, operation.target);
                break;
            }
            case 'insertRows': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                if (operation.row < 1 || operation.count < 1)
                    throw new Error(`invalid insertRows: row=${operation.row} count=${operation.count}`);
                sheet.spliceRows(operation.row, 0, ...Array.from({ length: operation.count }, () => []));
                shiftWorkbookRows(workbook, sheet.name, operation.row, operation.count);
                break;
            }
            case 'deleteRows': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                if (operation.row < 1 || operation.count < 1)
                    throw new Error(`invalid deleteRows: row=${operation.row} count=${operation.count}`);
                const end = operation.row + operation.count - 1;
                for (const formulaCell of collectDeletedRangeRefs(workbook, sheet.name, operation.row, end)) {
                    warnings.push({ op: index, message: `formula ${formulaCell} references a deleted row in ${sheet.name}` });
                }
                markDeletedRowRefs(workbook, sheet.name, operation.row, end);
                sheet.spliceRows(operation.row, operation.count);
                shiftWorkbookRows(workbook, sheet.name, end + 1, -operation.count);
                break;
            }
            case 'insertColumns': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const columnNumber = columnToNumber(operation.column);
                if (columnNumber < 1 || operation.count < 1)
                    throw new Error(`invalid insertColumns: column=${operation.column} count=${operation.count}`);
                sheet.spliceColumns(columnNumber, 0, ...Array.from({ length: operation.count }, () => []));
                shiftWorkbookColumns(workbook, sheet.name, columnNumber, operation.count);
                break;
            }
            case 'deleteColumns': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const columnNumber = columnToNumber(operation.column);
                if (columnNumber < 1 || operation.count < 1)
                    throw new Error(`invalid deleteColumns: column=${operation.column} count=${operation.count}`);
                const end = columnNumber + operation.count - 1;
                for (const formulaCell of collectDeletedColumnRefs(workbook, sheet.name, columnNumber, end)) {
                    warnings.push({ op: index, message: `formula ${formulaCell} references a deleted column in ${sheet.name}` });
                }
                markDeletedColumnRefs(workbook, sheet.name, columnNumber, end);
                sheet.spliceColumns(columnNumber, operation.count);
                shiftWorkbookColumns(workbook, sheet.name, end + 1, -operation.count);
                break;
            }
            case 'addSheet': {
                workbook.addWorksheet(operation.name);
                break;
            }
            case 'renameSheet': {
                const sheet = findSheet(workbook, operation.oldName);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.oldName}`);
                sheet.name = operation.newName;
                renameSheetReferences(workbook, operation.oldName, operation.newName);
                break;
            }
            case 'deleteSheet': {
                const sheet = findSheet(workbook, operation.name);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.name}`);
                workbook.removeWorksheet(sheet.id);
                break;
            }
            case 'clear': {
                for (const id of operation.cells)
                    resolveCell(workbook, id).value = null;
                break;
            }
            case 'merge': {
                applyMerge(workbook, operation.range, false);
                break;
            }
            case 'unmerge': {
                applyMerge(workbook, operation.range, true);
                break;
            }
            case 'copyRange': {
                copyRange(workbook, operation.source, operation.target, operation.move ?? false);
                break;
            }
            case 'fillSeries': {
                fillSeries(workbook, operation.start, operation.target, operation.step);
                break;
            }
            case 'style': {
                applyStyle(workbook, operation.range, operation.style);
                break;
            }
            case 'setColumnWidth': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                sheet.getColumn(operation.column).width = operation.width;
                break;
            }
            case 'setRowHeight': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                sheet.getRow(operation.row).height = operation.height;
                break;
            }
            case 'freezePanes': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const columnNumber = columnToNumber(operation.column);
                sheet.views = [{
                        state: 'frozen',
                        xSplit: Math.max(0, columnNumber - 1),
                        ySplit: Math.max(0, operation.row - 1),
                        topLeftCell: `${numberToColumn(columnNumber)}${operation.row}`,
                    }];
                break;
            }
            case 'findReplace': {
                const count = findReplace(workbook, operation.find, operation.replace, operation.sheet, operation.matchCase ?? false);
                warnings.push({ op: index, message: `findReplace replaced ${count} occurrence(s)` });
                break;
            }
            case 'duplicateSheet': {
                duplicateSheet(workbook, operation.name, operation.newName);
                break;
            }
            case 'hideSheet': {
                const sheet = findSheet(workbook, operation.name);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.name}`);
                sheet.state = operation.hidden === false ? 'visible' : 'hidden';
                break;
            }
            case 'setTabColor': {
                const sheet = findSheet(workbook, operation.name);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.name}`);
                sheet.properties.tabColor = { argb: normalizeColor(operation.color) };
                break;
            }
            case 'sortRange': {
                sortRange(workbook, operation.range, operation.keys, operation.headerRows ?? 0);
                warnings.push({ op: index, message: 'sortRange moved cell content; formulas outside the range still point to their original addresses' });
                break;
            }
            case 'dataValidation': {
                applyDataValidation(workbook, operation);
                break;
            }
            case 'conditionalFormatting': {
                applyConditionalFormatting(workbook, operation.range, operation.rules);
                break;
            }
            case 'autoFilter': {
                const parsed = parseRange(workbook, operation.range);
                parsed.sheet.autoFilter = {
                    from: { row: parsed.startRow, column: parsed.startCol },
                    to: { row: parsed.endRow, column: parsed.endCol },
                };
                break;
            }
            case 'addTable': {
                addTable(workbook, operation);
                break;
            }
        }
    });
    await workbook.xlsx.writeFile(outputPath);
    return { warnings };
}
function applyFill(workbook, sourceId, targetRange) {
    const source = resolveCell(workbook, sourceId);
    const sourceCell = parseCellId(sourceId);
    const bang = targetRange.lastIndexOf('!');
    const rawSheet = bang >= 0 ? targetRange.slice(0, bang) : null;
    const body = bang >= 0 ? targetRange.slice(bang + 1) : targetRange;
    const match = RANGE_LINE.exec(body);
    if (!match)
        throw new Error(`invalid fill target: ${targetRange}`);
    const targetSheetName = rawSheet ?? sourceCell.sheet;
    const sheet = findSheet(workbook, targetSheetName);
    if (!sheet)
        throw new Error(`sheet not found: ${targetSheetName}`);
    const startCol = columnToNumber(match[1]);
    const endCol = columnToNumber(match[3]);
    const startRow = Number(match[2]);
    const endRow = Number(match[4]);
    const content = cellContentOf(source);
    if (!content)
        return;
    for (let col = startCol; col <= endCol; col++) {
        for (let row = startRow; row <= endRow; row++) {
            if (col === columnToNumber(sourceCell.column) && row === sourceCell.row)
                continue;
            const cell = sheet.getCell(`${numberToColumn(col)}${row}`);
            const rowDelta = row - sourceCell.row;
            const colDelta = col - columnToNumber(sourceCell.column);
            const value = content.startsWith('=')
                ? shiftFormulaReferences(content, sourceCell.sheet, null, { rowDelta, colDelta })
                : content;
            writeContent(cell, value);
        }
    }
}
function shiftWorkbookRows(workbook, editedSheet, threshold, rowDelta) {
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
function collectDeletedRangeRefs(workbook, editedSheet, start, end) {
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
function shiftWorkbookColumns(workbook, editedSheet, threshold, colDelta) {
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
function collectDeletedColumnRefs(workbook, editedSheet, start, end) {
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
function markDeletedRowRefs(workbook, editedSheet, start, end) {
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
function markDeletedColumnRefs(workbook, editedSheet, start, end) {
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
function sortRange(workbook, range, keys, headerRows) {
    const parsed = parseRange(workbook, range);
    const keyColumns = keys.map((key) => ({
        column: columnToNumber(key.column),
        direction: key.direction ?? 'asc',
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
        const keyValues = [];
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const letter = numberToColumn(col);
            const cell = parsed.sheet.getCell(`${letter}${row}`);
            cells[letter] = cell.value;
            if (keyColumns.some((key) => key.column === col))
                keyValues.push(cell.value);
        }
        rows.push({ cells, keys: keyValues });
    }
    rows.sort((a, b) => compareSortRows(a.keys, b.keys, keyColumns.map((key) => key.direction)));
    for (let index = 0; index < rows.length; index++) {
        const targetRow = parsed.startRow + headerRows + index;
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const letter = numberToColumn(col);
            parsed.sheet.getCell(`${letter}${targetRow}`).value = rows[index].cells[letter] ?? null;
        }
    }
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
function applyDataValidation(workbook, options) {
    const parsed = parseRange(workbook, options.range);
    const validation = {
        type: options.type,
        operator: options.operator,
        formulae: [],
        allowBlank: options.allowBlank,
        showInputMessage: options.showInputMessage,
        prompt: options.prompt,
        showErrorMessage: options.showErrorMessage,
        errorStyle: options.errorStyle,
        errorTitle: options.errorTitle,
        error: options.error,
    };
    if (options.type === 'list') {
        if (!options.formula1)
            throw new Error('list data validation requires formula1 (comma-separated items or a range)');
        validation.formulae = [looksLikeRange(options.formula1) ? options.formula1 : `"${options.formula1}"`];
    }
    else if (options.formula1 !== undefined) {
        validation.formulae = [options.formula1];
        if (options.formula2 !== undefined)
            validation.formulae.push(options.formula2);
    }
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            parsed.sheet.getCell(`${numberToColumn(col)}${row}`).dataValidation = validation;
        }
    }
}
function looksLikeRange(value) {
    return /^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$/.test(value) || /[!$]/.test(value);
}
function applyConditionalFormatting(workbook, range, rules) {
    const parsed = parseRange(workbook, range);
    const mapped = rules.map((rule) => {
        const style = rule.style ? excelStyleToWorkbookStyle(rule.style) : undefined;
        if (rule.type === 'cellIs') {
            if (!rule.operator || rule.formula === undefined) {
                throw new Error('cellIs conditional formatting requires operator and formula');
            }
            return {
                type: 'cellIs',
                operator: rule.operator,
                formulae: [rule.formula, ...(rule.formula2 !== undefined ? [rule.formula2] : [])],
                style,
            };
        }
        return { type: 'expression', formulae: [String(rule.formula ?? '')], style };
    });
    const ref = `${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${parsed.endRow}`;
    parsed.sheet.addConditionalFormatting({ ref, rules: mapped });
}
function excelStyleToWorkbookStyle(style) {
    const result = {};
    if (style.bold !== undefined || style.italic !== undefined || style.underline !== undefined || style.fontColor !== undefined) {
        result.font = {
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            color: style.fontColor ? { argb: normalizeColor(style.fontColor) } : undefined,
        };
    }
    if (style.fill !== undefined) {
        result.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: normalizeColor(style.fill) } };
    }
    return result;
}
function addTable(workbook, options) {
    const parsed = parseRange(workbook, options.range);
    const ref = `${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${parsed.endRow}`;
    const headerRow = options.headerRow ?? true;
    const header = headerRow ? parsed.startRow : null;
    const dataStart = headerRow ? parsed.startRow + 1 : parsed.startRow;
    const columns = [];
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
        const letter = numberToColumn(col);
        const nameCell = header ? parsed.sheet.getCell(`${letter}${header}`).value : null;
        columns.push({ name: nameCell === null || nameCell === undefined ? `Column${letter}` : String(nameCell) });
    }
    const rows = [];
    for (let row = dataStart; row <= parsed.endRow; row++) {
        const values = [];
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            values.push(parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value);
        }
        rows.push(values);
    }
    parsed.sheet.addTable({
        name: options.name,
        ref,
        headerRow,
        totalsRow: options.totalsRow ?? false,
        columns,
        rows,
        style: {
            showRowStripes: options.showRowStripes ?? true,
            showColumnStripes: options.showColumnStripes ?? false,
        },
    });
}
function copyRange(workbook, sourceRange, targetCell, move) {
    const parsed = parseRange(workbook, sourceRange);
    const bang = targetCell.lastIndexOf('!');
    const targetSheetName = bang >= 0 ? targetCell.slice(0, bang) : parsed.sheet.name;
    const targetBody = bang >= 0 ? targetCell.slice(bang + 1) : targetCell;
    const match = /^([A-Za-z]{1,3})(\d+)$/.exec(targetBody);
    if (!match)
        throw new Error(`invalid target cell: ${targetCell}`);
    const targetSheet = findSheet(workbook, targetSheetName);
    if (!targetSheet)
        throw new Error(`sheet not found: ${targetSheetName}`);
    const targetCol = columnToNumber(match[1]);
    const targetRow = Number(match[2]);
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const source = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            const destCol = targetCol + (col - parsed.startCol);
            const destRow = targetRow + (row - parsed.startRow);
            const dest = targetSheet.getCell(`${numberToColumn(destCol)}${destRow}`);
            const content = cellContentOf(source);
            if (!content) {
                dest.value = null;
                continue;
            }
            dest.value = content.startsWith('=')
                ? {
                    formula: shiftFormulaReferences(content, parsed.sheet.name, null, {
                        rowDelta: destRow - row,
                        colDelta: destCol - col,
                    }).slice(1),
                }
                : toCellValue(content);
        }
    }
    if (move) {
        for (let row = parsed.startRow; row <= parsed.endRow; row++) {
            for (let col = parsed.startCol; col <= parsed.endCol; col++) {
                parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value = null;
            }
        }
    }
}
function fillSeries(workbook, startId, targetRange, step) {
    const startCell = resolveCell(workbook, startId);
    const startParsed = parseCellId(startId);
    const range = parseRange(workbook, targetRange);
    const startCol = columnToNumber(startParsed.column);
    if (startParsed.row !== range.startRow || startCol !== range.startCol) {
        throw new Error('fillSeries start cell must be the top-left cell of the target range');
    }
    const base = typeof startCell.value === 'number'
        ? startCell.value
        : startCell.value instanceof Date
            ? startCell.value.getTime()
            : null;
    if (base === null)
        throw new Error('fillSeries start cell must be a number or date');
    const isDate = startCell.value instanceof Date;
    const stepValue = step ?? (isDate ? 86_400_000 : 1);
    let index = 0;
    for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startCol; col <= range.endCol; col++) {
            if (row === startParsed.row && col === startCol)
                continue;
            index += 1;
            range.sheet.getCell(`${numberToColumn(col)}${row}`).value = isDate
                ? new Date(base + stepValue * index)
                : base + stepValue * index;
        }
    }
}
function applyStyle(workbook, range, style) {
    const parsed = parseRange(workbook, range);
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            const font = cell.font ?? {};
            if (style.bold !== undefined || style.italic !== undefined || style.underline !== undefined || style.fontColor !== undefined) {
                cell.font = {
                    ...font,
                    bold: style.bold ?? font.bold,
                    italic: style.italic ?? font.italic,
                    underline: style.underline ?? font.underline,
                    color: style.fontColor ? { argb: normalizeColor(style.fontColor) } : font.color,
                };
            }
            if (style.fill !== undefined) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: normalizeColor(style.fill) } };
            }
            if (style.numberFormat !== undefined)
                cell.numFmt = style.numberFormat;
            const alignment = cell.alignment ?? {};
            if (style.hAlign !== undefined || style.vAlign !== undefined || style.wrapText !== undefined) {
                cell.alignment = {
                    ...alignment,
                    horizontal: style.hAlign ?? alignment.horizontal,
                    vertical: style.vAlign ?? alignment.vertical,
                    wrapText: style.wrapText ?? alignment.wrapText,
                };
            }
        }
    }
}
function normalizeColor(color) {
    const hex = color.replace('#', '').trim();
    if (/^[0-9A-Fa-f]{6}$/.test(hex))
        return `FF${hex.toUpperCase()}`;
    if (/^[0-9A-Fa-f]{8}$/.test(hex))
        return hex.toUpperCase();
    throw new Error(`invalid color: ${color} (use 6-digit hex like FF0000)`);
}
function findReplace(workbook, find, replace, sheetName, matchCase) {
    let count = 0;
    const visit = (sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                const content = cellContentOf(cell);
                if (!content)
                    return;
                const replaced = replaceAllCase(content, find, replace, matchCase);
                if (replaced === content)
                    return;
                count += 1;
                cell.value = content.startsWith('=')
                    ? { formula: replaced.slice(1) }
                    : toCellValue(replaced);
            });
        });
    };
    if (sheetName) {
        const sheet = findSheet(workbook, sheetName);
        if (!sheet)
            throw new Error(`sheet not found: ${sheetName}`);
        visit(sheet);
    }
    else {
        workbook.eachSheet(visit);
    }
    return count;
}
function replaceAllCase(text, find, replace, matchCase) {
    if (matchCase)
        return text.replaceAll(find, replace);
    const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return text.replace(new RegExp(escaped, 'gi'), replace);
}
function duplicateSheet(workbook, name, newName) {
    const source = findSheet(workbook, name);
    if (!source)
        throw new Error(`sheet not found: ${name}`);
    if (findSheet(workbook, newName))
        throw new Error(`sheet already exists: ${newName}`);
    const copy = workbook.addWorksheet(newName);
    source.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
            copy.getCell(cell.address).value = cell.value;
        });
    });
    for (const merged of source.model.merges ?? [])
        copy.mergeCells(merged);
}
function renameSheetReferences(workbook, oldName, newName) {
    const oldQuoted = `'${oldName.replace(/'/g, "''")}'!`;
    const newQuoted = `'${newName.replace(/'/g, "''")}'!`;
    const oldBare = `${oldName}!`;
    const newBare = `${newName}!`;
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (!cell.formula)
                    return;
                const formula = cell.formula
                    .replaceAll(oldQuoted, newQuoted)
                    .replaceAll(oldBare, newBare);
                if (formula !== cell.formula)
                    cell.value = { formula };
            });
        });
    });
}
function applyMerge(workbook, range, unmerge) {
    const bang = range.lastIndexOf('!');
    const rawSheet = bang >= 0 ? range.slice(0, bang) : null;
    const body = bang >= 0 ? range.slice(bang + 1) : range;
    if (!rawSheet)
        throw new Error(`merge range requires a sheet: ${range}`);
    const sheet = findSheet(workbook, rawSheet);
    if (!sheet)
        throw new Error(`sheet not found: ${rawSheet}`);
    if (unmerge)
        sheet.unMergeCells(body);
    else
        sheet.mergeCells(body);
}
export async function operateWorkbookFile(path, operations, outputPath) {
    const result = await applyOperationsToWorkbook(path, operations, outputPath);
    const cells = await readWorkbookCells(await readFile(outputPath));
    const validation = validate(cells);
    return { ...result, outputPath, validation };
}
