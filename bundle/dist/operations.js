import ExcelJS from 'exceljs';
import { columnToNumber, normalizeSheet, numberToColumn, parseCellId, parseFormula, } from './formula.js';
import { guardFormulaInjection, parseCsv, stringifyCsv } from './csv.js';
import { validate } from './validator.js';
import { readWorkbookCells, stripPivotTableParts } from './workbook.js';
import { diffCellMaps, writePatchLog } from './diff.js';
import { annotateWorkbookXml, emptyAnnotations } from './xml-postprocess.js';
import { readFile, writeFile } from 'node:fs/promises';
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
/** Delete rows with the same reference-shift semantics as the deleteRows op. */
function deleteRowsFromSheet(workbook, sheetName, start, count, warnings, opIndex) {
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
function deleteColumnsFromSheet(workbook, sheetName, column, count, warnings, opIndex) {
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
function properCase(text) {
    return text.toLowerCase().replace(/(^|\s)(\S)/g, (_match, sep, char) => `${sep}${char.toUpperCase()}`);
}
/** Fullwidth ASCII/space/punctuation to halfwidth, then trim and collapse spaces. */
function normalizeTextValue(text) {
    let out = '';
    for (const char of text) {
        const code = char.charCodeAt(0);
        if (code >= 0xff01 && code <= 0xff5e)
            out += String.fromCharCode(code - 0xfee0);
        else if (char === '\u3000')
            out += ' ';
        else if (char === '\u2018' || char === '\u2019')
            out += "'";
        else if (char === '\u201c' || char === '\u201d')
            out += '"';
        else
            out += char;
    }
    return out.trim().replace(/\s+/g, ' ');
}
/**
 * Normalized similarity in [0, 1] for fuzzy matching: exact match is 1,
 * otherwise 1 minus the Levenshtein distance ratio over the longer string.
 * Callers normalize (trim/lowercase) before calling.
 */
function similarity(a, b) {
    if (a === b)
        return 1;
    if (a.length === 0 || b.length === 0)
        return 0;
    if (a.length > b.length)
        return similarity(b, a);
    const previous = Array.from({ length: a.length + 1 }, (_, i) => i);
    let current = new Array(a.length + 1);
    for (let j = 1; j <= b.length; j++) {
        current[0] = j;
        for (let i = 1; i <= a.length; i++) {
            const cost = a[i - 1] === b[j - 1] ? 0 : 1;
            current[i] = Math.min(previous[i] + 1, current[i - 1] + 1, previous[i - 1] + cost);
        }
        for (let i = 0; i <= a.length; i++)
            previous[i] = current[i];
    }
    return 1 - (previous[a.length] / b.length);
}
export async function applyOperationsToWorkbook(inputPath, operations, outputPath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(stripPivotTableParts(await readFile(inputPath)));
    const warnings = [];
    const annotations = emptyAnnotations();
    for (const [index, operation] of operations.entries()) {
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
                deleteRowsFromSheet(workbook, operation.sheet, operation.row, operation.count, warnings, index);
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
                deleteColumnsFromSheet(workbook, operation.sheet, columnToNumber(operation.column), operation.count, warnings, index);
                break;
            }
            case 'dedupeRows': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const columns = operation.columns && operation.columns.length > 0
                    ? operation.columns.map((column) => columnToNumber(column))
                    : Array.from({ length: sheet.columnCount }, (_, i) => i + 1);
                const keep = operation.keep ?? 'first';
                const rowsToDelete = [];
                const seen = new Set();
                const visit = (row) => {
                    const key = columns.map((col) => cellContentOf(sheet.getCell(`${numberToColumn(col)}${row}`))).join('\u0001');
                    if (seen.has(key))
                        rowsToDelete.push(row);
                    else
                        seen.add(key);
                };
                if (keep === 'first') {
                    for (let row = 1; row <= sheet.rowCount; row++)
                        visit(row);
                }
                else {
                    for (let row = sheet.rowCount; row >= 1; row--)
                        visit(row);
                }
                for (const row of rowsToDelete.sort((a, b) => b - a)) {
                    deleteRowsFromSheet(workbook, sheet.name, row, 1, warnings, index);
                }
                warnings.push({ op: index, message: `dedupeRows removed ${rowsToDelete.length} duplicate row(s) from ${sheet.name}` });
                break;
            }
            case 'fillMissing': {
                const parsed = parseRange(workbook, operation.range);
                if (operation.mode === 'value' && operation.value === undefined) {
                    throw new Error('value is required when fillMissing mode is "value"');
                }
                let filled = 0;
                for (let row = parsed.startRow; row <= parsed.endRow; row++) {
                    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
                        const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
                        if (cellContentOf(cell) !== '')
                            continue;
                        if (operation.mode === 'value') {
                            writeContent(cell, String(operation.value));
                            filled++;
                        }
                        else if (operation.mode === 'forward') {
                            for (let above = row - 1; above >= parsed.startRow; above--) {
                                const source = parsed.sheet.getCell(`${numberToColumn(col)}${above}`);
                                if (cellContentOf(source) === '')
                                    continue;
                                if (!source.formula)
                                    cell.value = source.value;
                                filled++;
                                break;
                            }
                        }
                        else {
                            for (let left = col - 1; left >= parsed.startCol; left--) {
                                const source = parsed.sheet.getCell(`${numberToColumn(left)}${row}`);
                                if (cellContentOf(source) === '')
                                    continue;
                                if (!source.formula)
                                    cell.value = source.value;
                                filled++;
                                break;
                            }
                        }
                    }
                }
                warnings.push({ op: index, message: `fillMissing filled ${filled} cell(s)` });
                break;
            }
            case 'removeEmptyRows': {
                const parsed = parseRange(workbook, operation.range);
                const emptyRows = [];
                for (let row = parsed.startRow; row <= parsed.endRow; row++) {
                    let empty = true;
                    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
                        if (cellContentOf(parsed.sheet.getCell(`${numberToColumn(col)}${row}`)) !== '') {
                            empty = false;
                            break;
                        }
                    }
                    if (empty)
                        emptyRows.push(row);
                }
                for (const row of emptyRows.sort((a, b) => b - a)) {
                    deleteRowsFromSheet(workbook, parsed.sheet.name, row, 1, warnings, index);
                }
                warnings.push({ op: index, message: `removeEmptyRows removed ${emptyRows.length} fully empty row(s) in ${operation.range}` });
                break;
            }
            case 'removeEmptyColumns': {
                const parsed = parseRange(workbook, operation.range);
                const emptyCols = [];
                for (let col = parsed.startCol; col <= parsed.endCol; col++) {
                    let empty = true;
                    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
                        if (cellContentOf(parsed.sheet.getCell(`${numberToColumn(col)}${row}`)) !== '') {
                            empty = false;
                            break;
                        }
                    }
                    if (empty)
                        emptyCols.push(col);
                }
                for (const col of emptyCols.sort((a, b) => b - a)) {
                    deleteColumnsFromSheet(workbook, parsed.sheet.name, col, 1, warnings, index);
                }
                warnings.push({ op: index, message: `removeEmptyColumns removed ${emptyCols.length} fully empty column(s) in ${operation.range}` });
                break;
            }
            case 'trimText': {
                const parsed = parseRange(workbook, operation.range);
                let trimmed = 0;
                for (let row = parsed.startRow; row <= parsed.endRow; row++) {
                    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
                        const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
                        if (cell.formula || typeof cell.value !== 'string')
                            continue;
                        const next = cell.value.trim();
                        if (next !== cell.value) {
                            cell.value = next;
                            trimmed++;
                        }
                    }
                }
                warnings.push({ op: index, message: `trimText trimmed ${trimmed} cell(s)` });
                break;
            }
            case 'changeCase': {
                const parsed = parseRange(workbook, operation.range);
                let changed = 0;
                const convert = (text) => operation.case === 'upper' ? text.toUpperCase() : operation.case === 'lower' ? text.toLowerCase() : properCase(text);
                for (let row = parsed.startRow; row <= parsed.endRow; row++) {
                    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
                        const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
                        if (cell.formula || typeof cell.value !== 'string')
                            continue;
                        const next = convert(cell.value);
                        if (next !== cell.value) {
                            cell.value = next;
                            changed++;
                        }
                    }
                }
                warnings.push({ op: index, message: `changeCase converted ${changed} cell(s) to ${operation.case}` });
                break;
            }
            case 'normalizeText': {
                const parsed = parseRange(workbook, operation.range);
                let normalized = 0;
                for (let row = parsed.startRow; row <= parsed.endRow; row++) {
                    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
                        const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
                        if (cell.formula || typeof cell.value !== 'string')
                            continue;
                        const next = normalizeTextValue(cell.value);
                        if (next !== cell.value) {
                            cell.value = next;
                            normalized++;
                        }
                    }
                }
                warnings.push({ op: index, message: `normalizeText normalized ${normalized} cell(s)` });
                break;
            }
            case 'splitColumn': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const columnNumber = columnToNumber(operation.column);
                const endRow = operation.endRow ?? sheet.rowCount;
                const partsByRow = new Map();
                let maxParts = 1;
                for (let row = operation.startRow; row <= endRow; row++) {
                    const text = cellContentOf(sheet.getCell(`${operation.column}${row}`));
                    if (!text)
                        continue;
                    const parts = text.split(operation.delimiter).map((part) => part.trim());
                    maxParts = Math.max(maxParts, parts.length);
                    partsByRow.set(row, parts);
                }
                if (maxParts > 1) {
                    sheet.spliceColumns(columnNumber + 1, 0, ...Array.from({ length: maxParts - 1 }, () => []));
                    shiftWorkbookColumns(workbook, sheet.name, columnNumber + 1, maxParts - 1);
                }
                for (const [row, parts] of partsByRow) {
                    for (let i = 0; i < maxParts; i++) {
                        // Split results are text fragments: preserve exactness (e.g. "01").
                        sheet.getCell(`${numberToColumn(columnNumber + i)}${row}`).value = parts[i] ?? '';
                    }
                }
                warnings.push({ op: index, message: `splitColumn split ${partsByRow.size} row(s) into up to ${maxParts} columns` });
                break;
            }
            case 'highlightRows': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const parsed = parseRange(workbook, operation.range);
                const style = operation.style ?? { fill: 'FFFF00' };
                let matched = 0;
                for (let row = parsed.startRow; row <= parsed.endRow; row++) {
                    let rowMatches = operation.criteria.every((criterion) => {
                        const cell = sheet.getCell(`${criterion.column}${row}`);
                        return matchesCriterion(cell.value, criterion.operator, criterion.value);
                    });
                    if (!rowMatches)
                        continue;
                    matched++;
                    applyStyle(workbook, `${sheet.name}!${numberToColumn(parsed.startCol)}${row}:${numberToColumn(parsed.endCol)}${row}`, style);
                }
                warnings.push({ op: index, message: `highlightRows highlighted ${matched} row(s) in ${operation.range}` });
                break;
            }
            case 'fuzzyMatch': {
                const sourceParsed = parseRange(workbook, operation.source);
                const targetParsed = parseRange(workbook, operation.target);
                const targetKeyCol = columnToNumber(operation.targetKey);
                const targetValueCol = columnToNumber(operation.valueColumn);
                const targetRows = [];
                for (let row = targetParsed.startRow; row <= targetParsed.endRow; row++) {
                    const key = cellContentOf(targetParsed.sheet.getCell(`${numberToColumn(targetKeyCol)}${row}`)).trim().toLowerCase();
                    if (!key)
                        continue;
                    targetRows.push({ key, value: cellContentOf(targetParsed.sheet.getCell(`${numberToColumn(targetValueCol)}${row}`)) });
                }
                const threshold = operation.threshold ?? 0.6;
                const outputCol = columnToNumber(operation.outputColumn);
                const scoreCol = operation.scoreColumn ? columnToNumber(operation.scoreColumn) : null;
                const sourceKeyCol = columnToNumber(operation.sourceKey);
                let matched = 0;
                for (let row = sourceParsed.startRow; row <= sourceParsed.endRow; row++) {
                    const key = cellContentOf(sourceParsed.sheet.getCell(`${numberToColumn(sourceKeyCol)}${row}`)).trim().toLowerCase();
                    if (!key)
                        continue;
                    let bestScore = 0;
                    let bestValue = '';
                    for (const target of targetRows) {
                        const score = similarity(key, target.key);
                        if (score > bestScore) {
                            bestScore = score;
                            bestValue = target.value;
                        }
                    }
                    if (bestScore >= threshold) {
                        matched++;
                        sourceParsed.sheet.getCell(`${numberToColumn(outputCol)}${row}`).value = bestValue;
                        if (scoreCol !== null)
                            sourceParsed.sheet.getCell(`${numberToColumn(scoreCol)}${row}`).value = Math.round(bestScore * 100) / 100;
                    }
                }
                warnings.push({ op: index, message: `fuzzyMatch matched ${matched}/${sourceParsed.endRow - sourceParsed.startRow + 1} source row(s) at threshold ${threshold}` });
                break;
            }
            case 'hideRows': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                if (operation.from < 1 || operation.to < operation.from)
                    throw new Error(`invalid hideRows: from=${operation.from} to=${operation.to}`);
                const hidden = operation.hidden ?? true;
                for (let row = operation.from; row <= operation.to; row++) {
                    const target = sheet.getRow(row);
                    target.hidden = hidden;
                    // ExcelJS drops empty rows on save unless they carry a height, so an
                    // empty hidden row needs one to survive the round-trip.
                    if (hidden && !target.hasValues && target.height === undefined) {
                        target.height = sheet.properties.defaultRowHeight ?? 15;
                    }
                }
                break;
            }
            case 'hideColumns': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                if (!operation.columns.length)
                    throw new Error('hideColumns requires at least one column');
                const hidden = operation.hidden ?? true;
                for (const column of operation.columns) {
                    sheet.getColumn(columnToNumber(column)).hidden = hidden;
                }
                break;
            }
            case 'groupRows': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const { start, end } = operation;
                const level = operation.level ?? 1;
                if (start < 1 || end < start)
                    throw new Error(`invalid groupRows: start=${start} end=${end}`);
                for (let row = start; row <= end; row++) {
                    const target = sheet.getRow(row);
                    target.outlineLevel = level;
                    if (operation.collapse) {
                        target.hidden = true;
                        if (!target.hasValues && target.height === undefined) {
                            target.height = sheet.properties.defaultRowHeight ?? 15;
                        }
                    }
                }
                sheet.properties.outlineLevelRow = Math.max(sheet.properties.outlineLevelRow ?? 0, level);
                break;
            }
            case 'groupColumns': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const from = columnToNumber(operation.from);
                const to = columnToNumber(operation.to);
                const level = operation.level ?? 1;
                if (from < 1 || to < from)
                    throw new Error(`invalid groupColumns: from=${operation.from} to=${operation.to}`);
                for (let col = from; col <= to; col++) {
                    const target = sheet.getColumn(col);
                    target.outlineLevel = level;
                    if (operation.collapse)
                        target.hidden = true;
                }
                sheet.properties.outlineLevelCol = Math.max(sheet.properties.outlineLevelCol ?? 0, level);
                break;
            }
            case 'autoFitColumnWidths': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const columns = operation.columns?.length
                    ? operation.columns.map((column) => columnToNumber(column))
                    : Array.from({ length: sheet.columnCount }, (_, i) => i + 1);
                const minWidth = operation.minWidth ?? 8;
                const maxWidth = operation.maxWidth ?? 60;
                for (const col of columns) {
                    let widest = 0;
                    for (let row = 1; row <= sheet.rowCount; row++) {
                        const text = displayTextOf(sheet.getCell(`${numberToColumn(col)}${row}`));
                        widest = Math.max(widest, displayWidth(text));
                    }
                    sheet.getColumn(col).width = Math.min(maxWidth, Math.max(minWidth, widest + 2));
                }
                break;
            }
            case 'unfreezePanes': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                sheet.views = [];
                break;
            }
            case 'transpose': {
                transposeRange(workbook, operation.source, operation.target);
                warnings.push({ op: index, message: 'transpose copied values and formulas (styles are not transposed)' });
                break;
            }
            case 'clearRange': {
                clearRange(workbook, operation.range, operation.mode ?? 'contents');
                break;
            }
            case 'joinSheets': {
                joinSheets(workbook, operation, warnings, index);
                break;
            }
            case 'crosstab': {
                applyCrosstab(workbook, operation, warnings, index);
                break;
            }
            case 'setHyperlink': {
                setHyperlink(workbook, operation);
                break;
            }
            case 'printTitles': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                if (operation.rows)
                    sheet.pageSetup.printTitlesRow = operation.rows;
                if (operation.columns)
                    sheet.pageSetup.printTitlesColumn = operation.columns;
                break;
            }
            case 'copyStyle': {
                copyStyle(workbook, operation.source, operation.target);
                break;
            }
            case 'freezeFormulas': {
                const frozen = freezeFormulas(workbook, operation.range);
                warnings.push({ op: index, message: `freezeFormulas converted ${frozen} formula(s) to their cached values` });
                break;
            }
            case 'uniqueValues': {
                const extracted = uniqueValues(workbook, operation);
                warnings.push({ op: index, message: `uniqueValues extracted ${extracted} distinct value(s)` });
                break;
            }
            case 'unmergeAll': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const merges = [...(sheet.model.merges ?? [])];
                for (const range of merges)
                    sheet.unMergeCells(range);
                warnings.push({ op: index, message: `unmergeAll removed ${merges.length} merged range(s) from ${sheet.name}` });
                break;
            }
            case 'setZoom': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                if (operation.zoom < 10 || operation.zoom > 400)
                    throw new Error(`invalid zoom: ${operation.zoom} (10-400)`);
                applySheetView(sheet, (view) => {
                    view.zoomScale = operation.zoom;
                    view.zoomScaleNormal = operation.normalZoom ?? operation.zoom;
                });
                break;
            }
            case 'showGridLines': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                applySheetView(sheet, (view) => {
                    view.showGridLines = operation.visible;
                });
                break;
            }
            case 'headerFooter': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                const headerFooter = {
                    ...sheet.headerFooter,
                    ...(operation.oddHeader !== undefined ? { oddHeader: operation.oddHeader } : {}),
                    ...(operation.oddFooter !== undefined ? { oddFooter: operation.oddFooter } : {}),
                    ...(operation.evenHeader !== undefined ? { evenHeader: operation.evenHeader } : {}),
                    ...(operation.evenFooter !== undefined ? { evenFooter: operation.evenFooter } : {}),
                    ...(operation.firstHeader !== undefined ? { firstHeader: operation.firstHeader } : {}),
                    ...(operation.firstFooter !== undefined ? { firstFooter: operation.firstFooter } : {}),
                    ...(operation.differentOddEven !== undefined ? { differentOddEven: operation.differentOddEven } : {}),
                    ...(operation.differentFirst !== undefined ? { differentFirst: operation.differentFirst } : {}),
                };
                sheet.headerFooter = headerFooter;
                break;
            }
            case 'moveSheet': {
                moveSheet(workbook, operation.name, operation.position);
                break;
            }
            case 'setWorkbookProperties': {
                if (operation.creator !== undefined)
                    workbook.creator = operation.creator;
                if (operation.lastModifiedBy !== undefined)
                    workbook.lastModifiedBy = operation.lastModifiedBy;
                if (operation.title !== undefined)
                    workbook.title = operation.title;
                if (operation.subject !== undefined)
                    workbook.subject = operation.subject;
                if (operation.description !== undefined)
                    workbook.description = operation.description;
                if (operation.keywords !== undefined)
                    workbook.keywords = operation.keywords;
                if (operation.recalcOnOpen)
                    workbook.calcProperties.fullCalcOnLoad = true;
                break;
            }
            case 'rankColumn': {
                applyRankColumn(workbook, operation);
                break;
            }
            case 'rowPageBreaks': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                if (!operation.rows.length)
                    throw new Error('rowPageBreaks requires at least one row');
                // OOXML <brk id> is zero-based: a break "above 1-based row N" is id=N-1.
                // ExcelJS models breaks on the worksheet (runtime property, untyped).
                const target = sheet;
                target.rowBreaks = operation.rows.map((row) => ({ id: row - 1, max: 16383, min: 0, man: true }));
                break;
            }
            case 'clearPageBreaks': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                sheet.rowBreaks = [];
                break;
            }
            case 'addComment': {
                const parsed = parseCellId(operation.cell);
                const sheet = findSheet(workbook, parsed.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${parsed.sheet}`);
                if (!operation.text.trim())
                    throw new Error('addComment requires non-empty text');
                const list = annotations.comments.get(sheet.name) ?? [];
                list.push({
                    ref: `${parsed.column}${parsed.row}`,
                    text: operation.text,
                    author: operation.author ?? 'dsh-excel-chat',
                    width: operation.width ?? 108,
                    height: operation.height ?? 60,
                });
                annotations.comments.set(sheet.name, list);
                break;
            }
            case 'addSparklines': {
                const dataBang = operation.dataRange.lastIndexOf('!');
                if (dataBang < 0)
                    throw new Error(`addSparklines dataRange requires a sheet: ${operation.dataRange}`);
                const sheetName = operation.dataRange.slice(0, dataBang);
                if (!findSheet(workbook, sheetName))
                    throw new Error(`sheet not found: ${sheetName}`);
                const groups = annotations.sparklines.get(sheetName) ?? [];
                groups.push({
                    dataRange: operation.dataRange,
                    locationRange: operation.locationRange,
                    type: operation.type ?? 'line',
                    color: normalizeColor(operation.color ?? '375623'),
                    negativeColor: normalizeColor(operation.negativeColor ?? 'D00000'),
                    markers: operation.markers ?? false,
                    highColor: normalizeColor(operation.highColor ?? 'FF7C00'),
                    lowColor: normalizeColor(operation.lowColor ?? 'D00000'),
                });
                annotations.sparklines.set(sheetName, groups);
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
                copyRange(workbook, operation.source, operation.target, operation.move ?? false, operation.valuesOnly ?? false);
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
            case 'importCsv': {
                await importCsv(workbook, operation);
                break;
            }
            case 'exportCsv': {
                await exportCsv(workbook, operation);
                break;
            }
            case 'sortRange': {
                sortRange(workbook, operation.range, operation.keys, operation.headerRows ?? 0);
                warnings.push({ op: index, message: 'sortRange moved cell content; formulas outside the range still point to their original addresses' });
                break;
            }
            case 'report': {
                applyReport(workbook, operation);
                break;
            }
            case 'preset': {
                applyPreset(workbook, operation);
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
            case 'subtotal': {
                applySubtotal(workbook, operation);
                warnings.push({ op: index, message: 'subtotal groups data by the group column; sort the range by that column first for correct grouping' });
                break;
            }
            case 'aggregateReport': {
                applyAggregateReport(workbook, operation);
                break;
            }
            case 'filterToRange': {
                applyFilterToRange(workbook, operation);
                break;
            }
            case 'protectSheet': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                sheet.protect(operation.password ?? '', {
                    selectLockedCells: operation.options?.selectLockedCells ?? true,
                    selectUnlockedCells: operation.options?.selectUnlockedCells ?? true,
                    formatCells: operation.options?.formatCells ?? false,
                    formatColumns: operation.options?.formatColumns ?? false,
                    formatRows: operation.options?.formatRows ?? false,
                    insertColumns: operation.options?.insertColumns ?? false,
                    insertRows: operation.options?.insertRows ?? false,
                    deleteColumns: operation.options?.deleteColumns ?? false,
                    deleteRows: operation.options?.deleteRows ?? false,
                    sort: operation.options?.sort ?? false,
                    autoFilter: operation.options?.autoFilter ?? false,
                });
                break;
            }
            case 'unprotectSheet': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                sheet.unprotect();
                break;
            }
            case 'mailMerge': {
                applyMailMerge(workbook, operation);
                break;
            }
            case 'pageSetup': {
                applyPageSetup(workbook, operation);
                break;
            }
            case 'definedName': {
                workbook.definedNames.add(operation.ref, operation.name);
                break;
            }
            case 'addTable': {
                addTable(workbook, operation);
                break;
            }
        }
    }
    const buffer = await workbook.xlsx.writeBuffer();
    if (annotations.comments.size > 0 || annotations.sparklines.size > 0) {
        // ExcelJS cannot write comments or sparklines; inject the XML parts now.
        const sheetFileOf = new Map();
        workbook.eachSheet((sheet) => {
            sheetFileOf.set(sheet.name, `xl/worksheets/sheet${sheet.id}.xml`);
        });
        const annotated = annotateWorkbookXml(new Uint8Array(buffer), annotations, sheetFileOf);
        await writeFile(outputPath, annotated);
    }
    else {
        await writeFile(outputPath, Buffer.from(buffer));
    }
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
        if (rule.type === 'containsText') {
            if (!rule.text)
                throw new Error('containsText conditional formatting requires text');
            return {
                type: 'containsText',
                operator: 'containsText',
                text: rule.text,
                formulae: [`NOT(ISERROR(SEARCH("${rule.text}",A1)))`],
                style,
            };
        }
        if (rule.type === 'notContainsText') {
            if (!rule.text)
                throw new Error('notContainsText conditional formatting requires text');
            return {
                type: 'expression',
                formulae: [`ISERROR(SEARCH("${rule.text}",A1))`],
                style,
            };
        }
        if (rule.type === 'blanks') {
            return { type: 'expression', formulae: ['ISBLANK(A1)'], style };
        }
        if (rule.type === 'noBlanks') {
            return { type: 'expression', formulae: ['NOT(ISBLANK(A1))'], style };
        }
        if (rule.type === 'errors') {
            return { type: 'expression', formulae: ['ISERROR(A1)'], style };
        }
        if (rule.type === 'noErrors') {
            return { type: 'expression', formulae: ['NOT(ISERROR(A1))'], style };
        }
        if (rule.type === 'duplicateValues' || rule.type === 'uniqueValues') {
            const range = `$${numberToColumn(parsed.startCol)}$${parsed.startRow}:$${numberToColumn(parsed.endCol)}$${parsed.endRow}`;
            const formula = rule.type === 'duplicateValues' ? `COUNTIF(${range},A1)>1` : `COUNTIF(${range},A1)=1`;
            return { type: 'expression', formulae: [formula], style };
        }
        if (rule.type === 'aboveAverage') {
            return { type: 'aboveAverage', style };
        }
        if (rule.type === 'belowAverage') {
            return { type: 'aboveAverage', aboveAverage: false, style };
        }
        if (rule.type === 'timePeriod') {
            return { type: 'timePeriod', timePeriod: rule.timePeriod ?? 'today', style };
        }
        if (rule.type === 'dataBar') {
            return {
                type: 'dataBar',
                color: { argb: normalizeColor(rule.color ?? '638EC6') },
                cfvo: [{ type: 'min' }, { type: 'max' }],
            };
        }
        if (rule.type === 'colorScale') {
            return {
                type: 'colorScale',
                cfvo: [
                    { type: 'min' },
                    { type: 'percentile', value: 50 },
                    { type: 'max' },
                ],
                color: [
                    { argb: normalizeColor(rule.minColor ?? 'F8696B') },
                    { argb: normalizeColor(rule.midColor ?? 'FFEB84') },
                    { argb: normalizeColor(rule.maxColor ?? '63BE7B') },
                ],
            };
        }
        if (rule.type === 'iconSet') {
            return {
                type: 'iconSet',
                iconSet: rule.iconSet ?? '3Arrows',
                cfvo: [
                    { type: 'percent', value: 0 },
                    { type: 'percent', value: 33 },
                    { type: 'percent', value: 67 },
                ],
            };
        }
        if (rule.type === 'top10') {
            return { type: 'top10', rank: rule.rank ?? 10, percent: rule.percent ?? false, bottom: rule.bottom ?? false };
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
const SUBTOTAL_CODES = {
    sum: 9,
    average: 1,
    count: 2,
    max: 4,
    min: 5,
};
function applySubtotal(workbook, options) {
    const parsed = parseRange(workbook, options.range);
    const groupCol = columnToNumber(options.groupColumn);
    if (groupCol < parsed.startCol || groupCol > parsed.endCol) {
        throw new Error(`subtotal group column outside range: ${options.groupColumn}`);
    }
    for (const summary of options.summaryColumns) {
        const col = columnToNumber(summary.column);
        if (col < parsed.startCol || col > parsed.endCol) {
            throw new Error(`subtotal summary column outside range: ${summary.column}`);
        }
        if (!SUBTOTAL_CODES[summary.function])
            throw new Error(`unsupported subtotal function: ${summary.function}`);
    }
    const sheet = parsed.sheet;
    const header = parsed.startRow;
    const firstData = parsed.startRow + 1;
    const lastData = parsed.endRow;
    const groups = [];
    let current = null;
    for (let row = firstData; row <= lastData; row++) {
        const raw = sheet.getCell(`${numberToColumn(groupCol)}${row}`).value;
        const key = raw === null || raw === undefined ? '' : String(raw);
        if (!current || current.value !== key) {
            current = { value: key, startRow: row, endRow: row };
            groups.push(current);
        }
        else {
            current.endRow = row;
        }
    }
    for (let groupIndex = 0; groupIndex < groups.length; groupIndex++) {
        const group = groups[groupIndex];
        const finalStartRow = group.startRow + groupIndex;
        const finalEndRow = group.endRow + groupIndex;
        const insertRow = finalEndRow + 1;
        sheet.spliceRows(insertRow, 0, []);
        shiftWorkbookRows(workbook, sheet.name, insertRow, 1);
        const label = sheet.getCell(`${numberToColumn(groupCol)}${insertRow}`);
        label.value = `${group.value} 汇总`;
        label.font = { bold: true };
        for (const summary of options.summaryColumns) {
            const col = columnToNumber(summary.column);
            const cell = sheet.getCell(`${numberToColumn(col)}${insertRow}`);
            cell.value = {
                formula: `SUBTOTAL(${SUBTOTAL_CODES[summary.function]},${numberToColumn(col)}${finalStartRow}:${numberToColumn(col)}${finalEndRow})`,
            };
            cell.font = { bold: true };
        }
    }
    if (options.addGrandTotal ?? true) {
        const totalRow = parsed.endRow + groups.length + 1;
        sheet.spliceRows(totalRow, 0, []);
        shiftWorkbookRows(workbook, sheet.name, totalRow, 1);
        const label = sheet.getCell(`${numberToColumn(groupCol)}${totalRow}`);
        label.value = '总计';
        label.font = { bold: true };
        for (const summary of options.summaryColumns) {
            const col = columnToNumber(summary.column);
            const cell = sheet.getCell(`${numberToColumn(col)}${totalRow}`);
            cell.value = {
                formula: `SUBTOTAL(${SUBTOTAL_CODES[summary.function]},${numberToColumn(col)}${firstData}:${numberToColumn(col)}${lastData + groups.length})`,
            };
            cell.font = { bold: true };
        }
    }
    void header;
    return groups.length + (options.addGrandTotal ?? true ? 1 : 0);
}
/**
 * One-shot report template: sort, subtotals, a dynamic SUMIFS summary sheet,
 * auto filter, header style, frozen header, and optional number format.
 * Ordering matters: subtotals run before the summary so its SUMIFS ranges
 * already cover the final data block (subtotal rows do not match group keys).
 */
function applyReport(workbook, options) {
    const parsed = parseRange(workbook, options.source);
    const sheet = parsed.sheet;
    const groupCol = columnToNumber(options.groupColumn);
    if (groupCol < parsed.startCol || groupCol > parsed.endCol) {
        throw new Error(`report group column outside range: ${options.groupColumn}`);
    }
    if (options.sort ?? true) {
        sortRange(workbook, options.source, [{ column: options.groupColumn }], 1);
    }
    let finalEndRow = parsed.endRow;
    if (options.subtotal ?? true) {
        const subtotalMetrics = options.metrics.map((metric) => ({
            column: metric.column,
            function: metric.function === 'counta' ? 'count' : metric.function,
        }));
        const inserted = applySubtotal(workbook, {
            op: 'subtotal',
            sheet: sheet.name,
            range: options.source,
            groupColumn: options.groupColumn,
            summaryColumns: subtotalMetrics,
            addGrandTotal: true,
        });
        finalEndRow = parsed.endRow + inserted;
    }
    const summarySheet = options.outputSheet ?? `${sheet.name}-汇总`;
    applyAggregateReport(workbook, {
        op: 'aggregateReport',
        source: `${sheet.name}!${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${finalEndRow}`,
        groupColumn: options.groupColumn,
        metrics: options.metrics,
        outputSheet: summarySheet,
    });
    if (options.autoFilter ?? true) {
        sheet.autoFilter = {
            from: { row: parsed.startRow, column: parsed.startCol },
            to: { row: finalEndRow, column: parsed.endCol },
        };
    }
    if (options.headerStyle ?? true) {
        applyStyle(workbook, `${sheet.name}!${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${parsed.startRow}`, {
            bold: true,
            fill: 'D9D9D9',
        });
    }
    if (options.freezeHeader ?? true) {
        sheet.views = [{
                state: 'frozen',
                xSplit: Math.max(0, parsed.startCol - 1),
                ySplit: Math.max(0, parsed.startRow),
                topLeftCell: `${numberToColumn(parsed.startCol)}${parsed.startRow + 1}`,
            }];
    }
    if (options.numberFormat) {
        for (const metric of options.metrics) {
            const col = columnToNumber(metric.column);
            for (let row = parsed.startRow; row <= finalEndRow; row++) {
                sheet.getCell(`${numberToColumn(col)}${row}`).numFmt = options.numberFormat;
            }
        }
        const summary = findSheet(workbook, summarySheet);
        if (summary) {
            options.metrics.forEach((metric, index) => {
                const col = numberToColumn(2 + index);
                for (let row = 1; row <= summary.rowCount; row++) {
                    summary.getCell(`${col}${row}`).numFmt = options.numberFormat;
                }
            });
        }
    }
}
const ROLE_LABELS = {
    ops: '运营报表',
    product: '产品分析',
    data: '数据分析',
};
/**
 * Role-based one-shot preset: 运营 gets a report with data bars, 产品 and 数分
 * get a report with color scales, and 数分 additionally writes a filtered copy.
 */
function applyPreset(workbook, options) {
    const parsed = parseRange(workbook, options.source);
    const sheet = parsed.sheet;
    const summarySheet = `${sheet.name}-${ROLE_LABELS[options.role]}`;
    if (options.filter) {
        const filterSheetName = `${sheet.name}-筛选`;
        if (!findSheet(workbook, filterSheetName))
            workbook.addWorksheet(filterSheetName);
        applyFilterToRange(workbook, {
            op: 'filterToRange',
            source: options.source,
            criteria: [options.filter],
            target: `${filterSheetName}!A1`,
        });
    }
    applyReport(workbook, {
        op: 'report',
        source: options.source,
        groupColumn: options.groupColumn,
        metrics: options.metrics,
        numberFormat: '#,##0.00',
        outputSheet: summarySheet,
    });
    for (const metric of options.metrics) {
        const col = numberToColumn(columnToNumber(metric.column));
        const range = `${sheet.name}!${col}${parsed.startRow}:${col}${sheet.rowCount}`;
        if (options.role === 'ops') {
            applyConditionalFormatting(workbook, range, [{ type: 'dataBar', color: '63BE7B' }]);
        }
        else {
            applyConditionalFormatting(workbook, range, [{
                    type: 'colorScale',
                    minColor: 'F8696B',
                    midColor: 'FFEB84',
                    maxColor: '63BE7B',
                }]);
        }
    }
}
const REPORT_FUNCTIONS = {
    sum: 'SUMIFS',
    average: 'AVERAGEIFS',
    count: 'COUNTIFS',
    counta: 'COUNTIFS',
    max: 'MAXIFS',
    min: 'MINIFS',
};
function applyAggregateReport(workbook, options) {
    const parsed = parseRange(workbook, options.source);
    const groupCol = columnToNumber(options.groupColumn);
    const sourceSheet = parsed.sheet.name;
    const firstData = parsed.startRow + 1;
    const lastData = parsed.endRow;
    const groupRange = `${sourceSheet}!$${numberToColumn(groupCol)}$${firstData}:$${numberToColumn(groupCol)}$${lastData}`;
    const groupValues = [];
    const seen = new Set();
    for (let row = firstData; row <= lastData; row++) {
        const raw = parsed.sheet.getCell(`${numberToColumn(groupCol)}${row}`).value;
        const key = raw === null || raw === undefined ? '' : String(raw);
        if (!seen.has(key)) {
            seen.add(key);
            groupValues.push(key);
        }
    }
    const outputSheetName = options.outputSheet ?? `${sourceSheet}-汇总`;
    let output = findSheet(workbook, outputSheetName);
    if (!output)
        output = workbook.addWorksheet(outputSheetName);
    const groupHeader = String(parsed.sheet.getCell(`${numberToColumn(groupCol)}${parsed.startRow}`).value ?? options.groupColumn);
    output.getCell('A1').value = groupHeader;
    output.getCell('A1').font = { bold: true };
    const metricLabels = {
        sum: '合计',
        average: '平均',
        count: '计数',
        counta: '非空计数',
        max: '最大',
        min: '最小',
    };
    options.metrics.forEach((metric, index) => {
        const metricCol = columnToNumber(metric.column);
        const header = String(parsed.sheet.getCell(`${numberToColumn(metricCol)}${parsed.startRow}`).value ?? metric.column);
        const cell = output.getCell(`${numberToColumn(2 + index)}1`);
        cell.value = `${header} ${metricLabels[metric.function]}`;
        cell.font = { bold: true };
        void metricCol;
    });
    for (let index = 0; index < groupValues.length; index++) {
        const row = 2 + index;
        const groupCell = output.getCell(`A${row}`);
        groupCell.value = groupValues[index];
        options.metrics.forEach((metric, metricIndex) => {
            const metricCol = columnToNumber(metric.column);
            const metricRange = `${sourceSheet}!$${numberToColumn(metricCol)}$${firstData}:$${numberToColumn(metricCol)}$${lastData}`;
            const fn = REPORT_FUNCTIONS[metric.function];
            const criteria = `A${row}`;
            output.getCell(`${numberToColumn(2 + metricIndex)}${row}`).value = {
                formula: `${fn}(${metricRange},${groupRange},${criteria})`,
            };
        });
    }
    const lastGroupRow = 1 + groupValues.length;
    options.metrics.forEach((metric, metricIndex) => {
        const col = numberToColumn(2 + metricIndex);
        output.getCell(`${col}${lastGroupRow + 1}`).value = {
            formula: `SUM(${col}2:${col}${lastGroupRow})`,
        };
        output.getCell(`${col}${lastGroupRow + 1}`).font = { bold: true };
    });
    output.getCell(`A${lastGroupRow + 1}`).value = '总计';
    output.getCell(`A${lastGroupRow + 1}`).font = { bold: true };
}
function applyFilterToRange(workbook, options) {
    const parsed = parseRange(workbook, options.source);
    const target = parseTargetCell(workbook, options.target, parsed.sheet.name);
    const matchAll = options.matchAll ?? true;
    const headerRow = [];
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
        headerRow.push(parsed.sheet.getCell(`${numberToColumn(col)}${parsed.startRow}`).value);
    }
    let targetRow = target.row;
    headerRow.forEach((value, index) => {
        target.sheet.getCell(`${numberToColumn(target.col + index)}${targetRow}`).value = value;
    });
    targetRow += 1;
    for (let row = parsed.startRow + 1; row <= parsed.endRow; row++) {
        let matched = matchAll;
        for (const criterion of options.criteria) {
            const col = columnToNumber(criterion.column);
            const actual = parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value;
            const ok = matchesCriterion(actual, criterion.operator, criterion.value);
            if (matchAll && !ok) {
                matched = false;
                break;
            }
            if (!matchAll && ok) {
                matched = true;
                break;
            }
        }
        if (!matched)
            continue;
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            target.sheet.getCell(`${numberToColumn(target.col + (col - parsed.startCol))}${targetRow}`).value =
                parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value;
        }
        targetRow += 1;
    }
}
function matchesCriterion(actual, operator, expected) {
    const actualNumber = typeof actual === 'number' ? actual : null;
    const expectedNumber = typeof expected === 'number' ? expected : Number(expected);
    const actualText = actual === null || actual === undefined ? '' : String(actual);
    const expectedText = String(expected);
    switch (operator) {
        case 'eq':
            return actualNumber !== null && Number.isFinite(expectedNumber)
                ? actualNumber === expectedNumber
                : actualText.toLowerCase() === expectedText.toLowerCase();
        case 'neq':
            return !matchesCriterion(actual, 'eq', expected);
        case 'contains':
            return actualText.toLowerCase().includes(expectedText.toLowerCase());
        case 'gt':
            return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber > expectedNumber;
        case 'gte':
            return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber >= expectedNumber;
        case 'lt':
            return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber < expectedNumber;
        case 'lte':
            return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber <= expectedNumber;
    }
}
function applyMailMerge(workbook, options) {
    const template = parseRange(workbook, options.template);
    const data = parseRange(workbook, options.data);
    const headers = new Map();
    for (let col = data.startCol; col <= data.endCol; col++) {
        const raw = data.sheet.getCell(`${numberToColumn(col)}${data.startRow}`).value;
        headers.set(String(raw ?? '').toLowerCase(), col);
    }
    const templateRows = [];
    for (let row = template.startRow; row <= template.endRow; row++) {
        const cells = {};
        for (let col = template.startCol; col <= template.endCol; col++) {
            cells[col] = template.sheet.getCell(`${numberToColumn(col)}${row}`).value;
        }
        templateRows.push(cells);
    }
    const outputSheetName = options.outputSheet ?? `${template.sheet.name}-合并`;
    let output = findSheet(workbook, outputSheetName);
    if (!output)
        output = workbook.addWorksheet(outputSheetName);
    let outputRow = 1;
    const placeholder = /\{([^{}]+)\}/g;
    for (let dataRow = data.startRow + 1; dataRow <= data.endRow; dataRow++) {
        const record = new Map();
        for (const [header, col] of headers) {
            record.set(header, data.sheet.getCell(`${numberToColumn(col)}${dataRow}`).value);
        }
        for (const templateRow of templateRows) {
            for (const [col, value] of Object.entries(templateRow)) {
                const column = Number(col);
                const text = value === null || value === undefined ? '' : String(value);
                if (/^\{[^{}]+\}$/.test(text.trim())) {
                    const key = text.trim().slice(1, -1).toLowerCase();
                    output.getCell(`${numberToColumn(column)}${outputRow}`).value = record.get(key) ?? text;
                    continue;
                }
                if (placeholder.test(text)) {
                    placeholder.lastIndex = 0;
                    output.getCell(`${numberToColumn(column)}${outputRow}`).value = text.replace(placeholder, (_match, name) => {
                        const replacement = record.get(String(name).toLowerCase());
                        return replacement === undefined ? _match : String(replacement);
                    });
                    continue;
                }
                output.getCell(`${numberToColumn(column)}${outputRow}`).value = value;
            }
            outputRow += 1;
        }
    }
}
function parseTargetCell(workbook, target, defaultSheet) {
    const bang = target.lastIndexOf('!');
    const sheetName = bang >= 0 ? target.slice(0, bang) : defaultSheet;
    const body = bang >= 0 ? target.slice(bang + 1) : target;
    const match = /^([A-Za-z]{1,3})(\d+)$/.exec(body);
    if (!match)
        throw new Error(`invalid target cell: ${target}`);
    const sheet = findSheet(workbook, sheetName);
    if (!sheet)
        throw new Error(`sheet not found: ${sheetName}`);
    return { sheet, col: columnToNumber(match[1]), row: Number(match[2]) };
}
function copyRange(workbook, sourceRange, targetCell, move, valuesOnly = false) {
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
            if (valuesOnly) {
                // Paste-special: values only. Formulas contribute their last cached
                // result; empty cells clear the destination.
                dest.value = source.formula ? (source.result ?? null) : source.value;
                continue;
            }
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
    const startContent = cellContentOf(startCell);
    if (startContent.startsWith('=')) {
        for (let row = range.startRow; row <= range.endRow; row++) {
            for (let col = range.startCol; col <= range.endCol; col++) {
                if (row === startParsed.row && col === startCol)
                    continue;
                const cell = range.sheet.getCell(`${numberToColumn(col)}${row}`);
                const shifted = shiftFormulaReferences(startContent, startParsed.sheet, null, {
                    rowDelta: row - startParsed.row,
                    colDelta: col - startCol,
                });
                writeContent(cell, shifted);
            }
        }
        return;
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
            if (style.bold !== undefined ||
                style.italic !== undefined ||
                style.underline !== undefined ||
                style.strikeThrough !== undefined ||
                style.fontColor !== undefined ||
                style.fontSize !== undefined ||
                style.fontName !== undefined) {
                cell.font = {
                    ...font,
                    bold: style.bold ?? font.bold,
                    italic: style.italic ?? font.italic,
                    underline: style.underline ?? font.underline,
                    strike: style.strikeThrough ?? font.strike,
                    size: style.fontSize ?? font.size,
                    name: style.fontName ?? font.name,
                    color: style.fontColor ? { argb: normalizeColor(style.fontColor) } : font.color,
                };
            }
            if (style.fill !== undefined) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: normalizeColor(style.fill) } };
            }
            if (style.numberFormat !== undefined)
                cell.numFmt = style.numberFormat;
            const alignment = cell.alignment ?? {};
            if (style.hAlign !== undefined ||
                style.vAlign !== undefined ||
                style.wrapText !== undefined ||
                style.textRotation !== undefined ||
                style.shrinkToFit !== undefined ||
                style.indent !== undefined) {
                cell.alignment = {
                    ...alignment,
                    horizontal: style.hAlign ?? alignment.horizontal,
                    vertical: style.vAlign ?? alignment.vertical,
                    wrapText: style.wrapText ?? alignment.wrapText,
                    textRotation: style.textRotation ?? alignment.textRotation,
                    shrinkToFit: style.shrinkToFit ?? alignment.shrinkToFit,
                    indent: style.indent ?? alignment.indent,
                };
            }
            if (style.border) {
                const border = {};
                for (const side of ['top', 'bottom', 'left', 'right']) {
                    const edge = style.border[side];
                    if (edge) {
                        border[side] = {
                            style: edge.style ?? 'thin',
                            color: edge.color ? { argb: normalizeColor(edge.color) } : undefined,
                        };
                    }
                }
                cell.border = border;
            }
        }
    }
}
function applyPageSetup(workbook, options) {
    const sheet = findSheet(workbook, options.sheet);
    if (!sheet)
        throw new Error(`sheet not found: ${options.sheet}`);
    const pageSetup = sheet.pageSetup;
    if (options.printArea)
        pageSetup.printArea = options.printArea;
    if (options.orientation)
        pageSetup.orientation = options.orientation;
    if (options.fitToPage !== undefined)
        pageSetup.fitToPage = options.fitToPage;
    if (options.fitToWidth !== undefined)
        pageSetup.fitToWidth = options.fitToWidth;
    if (options.fitToHeight !== undefined)
        pageSetup.fitToHeight = options.fitToHeight;
    if (options.margins)
        pageSetup.margins = { ...pageSetup.margins, ...options.margins };
    if (options.centerHorizontally !== undefined)
        pageSetup.horizontalCentered = options.centerHorizontally;
    if (options.centerVertically !== undefined)
        pageSetup.verticalCentered = options.centerVertically;
}
async function importCsv(workbook, options) {
    const text = await readFile(options.file, 'utf8');
    const rows = parseCsv(text, options.delimiter ?? ',');
    const sheetName = options.sheet ?? 'CSV';
    let sheet = findSheet(workbook, sheetName);
    if (!sheet)
        sheet = workbook.addWorksheet(sheetName);
    rows.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
            writeContent(sheet.getCell(`${numberToColumn(colIndex + 1)}${rowIndex + 1}`), value);
        });
    });
}
async function exportCsv(workbook, options) {
    const sheet = findSheet(workbook, options.sheet ?? workbook.worksheets[0].name);
    if (!sheet)
        throw new Error(`sheet not found: ${options.sheet}`);
    const parsed = options.range ? parseRange(workbook, `${sheet.name}!${options.range}`) : null;
    const startCol = parsed?.startCol ?? 1;
    const startRow = parsed?.startRow ?? 1;
    const endCol = parsed?.endCol ?? sheet.columnCount;
    const endRow = parsed?.endRow ?? sheet.rowCount;
    const guard = options.guardFormulas ?? true;
    const rows = [];
    for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
        const row = [];
        for (let colIndex = startCol; colIndex <= endCol; colIndex++) {
            const cell = sheet.getCell(`${numberToColumn(colIndex)}${rowIndex}`);
            if (cell.formula) {
                row.push(`=${cell.formula}`);
            }
            else {
                const raw = cell.value;
                let text = raw === null || raw === undefined ? '' : String(raw);
                if (guard && typeof raw === 'string')
                    text = guardFormulaInjection(text);
                row.push(text);
            }
        }
        rows.push(row);
    }
    await writeFile(options.file, stringifyCsv(rows, options.delimiter ?? ','), 'utf8');
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
/** Visible text of a cell for width estimation: formula cells use their cached result. */
function displayTextOf(cell) {
    const value = cell.formula ? cell.result : cell.value;
    if (value === null || value === undefined)
        return '';
    if (value instanceof Date)
        return '2026-12-31';
    if (typeof value === 'object')
        return JSON.stringify(value);
    return String(value);
}
/**
 * Approximate display width in character units: CJK/fullwidth characters count
 * as 2 columns, everything else as 1.
 */
function displayWidth(text) {
    let width = 0;
    for (const char of text) {
        const code = char.codePointAt(0) ?? 0;
        width += code > 0x2e7f ? 2 : 1;
    }
    return width;
}
function transposeRange(workbook, sourceRange, targetCell) {
    const parsed = parseRange(workbook, sourceRange);
    const target = parseTargetCell(workbook, targetCell, parsed.sheet.name);
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const source = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            // (row,col) maps to (targetRow + colOffset, targetCol + rowOffset).
            const destRow = target.row + (col - parsed.startCol);
            const destCol = target.col + (row - parsed.startRow);
            const dest = target.sheet.getCell(`${numberToColumn(destCol)}${destRow}`);
            const content = cellContentOf(source);
            if (!content)
                continue;
            dest.value = content.startsWith('=')
                ? toCellValue(shiftFormulaReferences(content, parsed.sheet.name, null, {
                    rowDelta: destRow - row,
                    colDelta: destCol - col,
                }))
                : source.value;
        }
    }
}
function clearRange(workbook, range, mode) {
    const parsed = parseRange(workbook, range);
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            if (mode === 'contents') {
                cell.value = null;
            }
            else if (mode === 'formats') {
                cell.style = {};
            }
            else {
                cell.value = null;
                cell.style = {};
            }
        }
    }
}
function joinSheets(workbook, operation, warnings, opIndex) {
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
                value.startsWith('=') ? { formula: value.slice(1) } : toCellValue(value);
        });
    }
    warnings.push({ op: opIndex, message: `joinSheets matched ${matched} row(s), ${missed} without a lookup hit` });
}
/** Join keys are compared trimmed + lowercased, numbers via their text form. */
function normalizeJoinKey(value) {
    if (value === null || value === undefined)
        return '';
    return String(typeof value === 'object' && !(value instanceof Date) ? JSON.stringify(value) : value).trim().toLowerCase();
}
const CROSSTAB_FUNCTIONS = {
    sum: 'SUMIFS',
    average: 'AVERAGEIFS',
    count: 'COUNTIFS',
    counta: 'COUNTIFS',
    max: 'MAXIFS',
    min: 'MINIFS',
};
/** Aggregations where a grand total of the computed grid is meaningful. */
const CROSSTAB_TOTALABLE = new Set(['sum', 'count', 'counta']);
function applyCrosstab(workbook, options, warnings, opIndex) {
    const needsMetric = options.metric.function !== 'count' && options.metric.function !== 'counta';
    if (needsMetric && !options.metric.column) {
        throw new Error(`crosstab function "${options.metric.function}" requires metric.column`);
    }
    const parsed = parseRange(workbook, options.source);
    const rowCol = columnToNumber(options.rowColumn);
    const colCol = columnToNumber(options.columnColumn);
    const firstData = parsed.startRow + 1;
    const collectKeys = (col) => {
        const keys = [];
        const seen = new Set();
        for (let row = firstData; row <= parsed.endRow; row++) {
            const raw = parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value;
            const key = raw === null || raw === undefined ? '' : String(raw);
            if (!seen.has(key)) {
                seen.add(key);
                keys.push(key);
            }
        }
        return keys;
    };
    const rowKeys = collectKeys(rowCol);
    const colKeys = collectKeys(colCol);
    if (!rowKeys.length || !colKeys.length)
        throw new Error('crosstab source has no data rows');
    const sheetRange = (col) => `${parsed.sheet.name}!$${numberToColumn(col)}$${firstData}:$${numberToColumn(col)}$${parsed.endRow}`;
    const rowRange = sheetRange(rowCol);
    const colRange = sheetRange(colCol);
    const metricRange = options.metric.column ? sheetRange(columnToNumber(options.metric.column)) : null;
    const fn = CROSSTAB_FUNCTIONS[options.metric.function];
    const outputSheetName = options.outputSheet ?? `${parsed.sheet.name}-交叉表`;
    let output = findSheet(workbook, outputSheetName);
    if (!output)
        output = workbook.addWorksheet(outputSheetName);
    const rowHeader = String(parsed.sheet.getCell(`${numberToColumn(rowCol)}${parsed.startRow}`).value ?? options.rowColumn);
    const colHeader = String(parsed.sheet.getCell(`${numberToColumn(colCol)}${parsed.startRow}`).value ?? options.columnColumn);
    const corner = output.getCell('A1');
    corner.value = `${rowHeader}\\${colHeader}`;
    corner.font = { bold: true };
    colKeys.forEach((key, i) => {
        const cell = output.getCell(`${numberToColumn(2 + i)}1`);
        cell.value = key;
        cell.font = { bold: true };
    });
    rowKeys.forEach((rowKey, rowIndex) => {
        const outRow = 2 + rowIndex;
        output.getCell(`A${outRow}`).value = rowKey;
        colKeys.forEach((_colKey, colIndex) => {
            const columnLetter = numberToColumn(2 + colIndex);
            // Criteria point at output-sheet cells, so keys never need quoting.
            const body = metricRange
                ? `${metricRange},${rowRange},$A${outRow},${colRange},${columnLetter}$1`
                : `${rowRange},$A${outRow},${colRange},${columnLetter}$1`;
            const formula = `${fn}(${body})`;
            output.getCell(`${columnLetter}${outRow}`).value = {
                formula: options.metric.function === 'average' ? `IFERROR(${formula},0)` : formula,
            };
        });
    });
    const totals = options.totals ?? true;
    if (totals && CROSSTAB_TOTALABLE.has(options.metric.function)) {
        const totalRow = 2 + rowKeys.length;
        const totalCol = 2 + colKeys.length;
        output.getCell(`A${totalRow}`).value = '总计';
        output.getCell(`A${totalRow}`).font = { bold: true };
        colKeys.forEach((_colKey, colIndex) => {
            const columnLetter = numberToColumn(2 + colIndex);
            const cell = output.getCell(`${columnLetter}${totalRow}`);
            cell.value = { formula: `SUM(${columnLetter}2:${columnLetter}${totalRow - 1})` };
            cell.font = { bold: true };
        });
        for (let row = 2; row <= totalRow; row++) {
            const last = numberToColumn(totalCol - 1);
            const cell = output.getCell(`${numberToColumn(totalCol)}${row}`);
            cell.value = { formula: `SUM(B${row}:${last}${row})` };
            if (row === totalRow)
                cell.font = { bold: true };
        }
    }
    warnings.push({
        op: opIndex,
        message: `crosstab built ${rowKeys.length}x${colKeys.length} grid on ${outputSheetName} with live ${fn} formulas`,
    });
}
function setHyperlink(workbook, options) {
    const cell = resolveCell(workbook, options.cell);
    const text = options.text;
    if (options.url) {
        cell.value = { text: text ?? options.url, hyperlink: options.url };
        return;
    }
    if (options.location) {
        const location = options.location.startsWith('#') ? options.location : `#${options.location}`;
        cell.value = { text: text ?? location.slice(1), hyperlink: location };
        return;
    }
    throw new Error('setHyperlink requires url (external) or location (internal, e.g. "Sheet2!A1")');
}
/** Clone font/fill/border/alignment/number format from one cell onto every cell in the target range. */
function copyStyle(workbook, sourceId, targetRange) {
    const source = resolveCell(workbook, sourceId);
    const style = JSON.parse(JSON.stringify(source.style ?? {}));
    const parsed = parseRange(workbook, targetRange);
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            cell.style = JSON.parse(JSON.stringify(style));
        }
    }
}
/** Replace formulas with their cached results ("paste values" in place). */
function freezeFormulas(workbook, range) {
    const parsed = parseRange(workbook, range);
    let frozen = 0;
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            if (!cell.formula)
                continue;
            const result = cell.result;
            cell.value = result === undefined || result === null ? null : result;
            frozen++;
        }
    }
    return frozen;
}
/** Write the distinct values of a source column into a target column, first-seen order. */
function uniqueValues(workbook, options) {
    const parsed = parseRange(workbook, options.source);
    const target = parseTargetCell(workbook, options.target, parsed.sheet.name);
    const seen = new Set();
    const ordered = [];
    const firstDataRow = options.includeHeader ? parsed.startRow : parsed.startRow + 1;
    for (let row = firstDataRow; row <= parsed.endRow; row++) {
        const cell = parsed.sheet.getCell(`${numberToColumn(parsed.startCol)}${row}`);
        const raw = cell.formula ? cell.result : cell.value;
        const key = raw === null || raw === undefined ? '' : String(raw);
        if (seen.has(key))
            continue;
        seen.add(key);
        ordered.push(raw);
    }
    let outRow = target.row;
    for (const value of ordered) {
        target.sheet.getCell(`${numberToColumn(target.col)}${outRow}`).value = value === null || value === undefined ? '' : value;
        outRow++;
    }
    if (options.includeHeader) {
        const header = parsed.sheet.getCell(`${numberToColumn(parsed.startCol)}${parsed.startRow}`).value;
        target.sheet.getCell(`${numberToColumn(target.col)}${target.row}`).value = header;
    }
    return ordered.length;
}
/** Patch every existing sheet view without dropping frozen panes or other flags. */
function applySheetView(sheet, patch) {
    // exceljs types views as strict unions but accepts partial views at runtime,
    // so the default view is cast from a minimal object.
    const existing = (sheet.views ?? []);
    const views = existing.length
        ? existing
        : [{ workbookViewId: 0 }];
    for (const view of views)
        patch(view);
    sheet.views = views;
}
/** Reorder sheets by rewriting orderNo (worksheets getter sorts by it). */
function moveSheet(workbook, name, position) {
    const sheet = findSheet(workbook, name);
    if (!sheet)
        throw new Error(`sheet not found: ${name}`);
    const ordered = workbook.worksheets;
    const others = ordered.filter((entry) => entry.id !== sheet.id);
    const clamped = Math.max(1, Math.min(position, ordered.length));
    const before = others.slice(0, clamped - 1);
    const after = others.slice(clamped - 1);
    [...before, sheet, ...after].forEach((entry, i) => {
        ;
        entry.orderNo = i + 1;
    });
}
/** Append a live RANK column next to a metric column. */
function applyRankColumn(workbook, options) {
    const parsed = parseRange(workbook, options.range);
    const metricCol = columnToNumber(options.metricColumn);
    if (metricCol < parsed.startCol || metricCol > parsed.endCol) {
        throw new Error(`rankColumn metric column outside range: ${options.metricColumn}`);
    }
    const firstData = options.skipHeader === false ? parsed.startRow : parsed.startRow + 1;
    const metricRange = `${parsed.sheet.name}!$${numberToColumn(metricCol)}$${firstData}:$${numberToColumn(metricCol)}$${parsed.endRow}`;
    for (let row = firstData; row <= parsed.endRow; row++) {
        parsed.sheet.getCell(`${options.outputColumn}${row}`).value = {
            formula: `RANK(${numberToColumn(metricCol)}${row},${metricRange},${options.descending === false ? 1 : 0})`,
        };
    }
}
export async function operateWorkbookFile(path, operations, outputPath) {
    const result = await applyOperationsToWorkbook(path, operations, outputPath);
    const [before, after] = await Promise.all([
        readWorkbookCells(await readFile(path)),
        readWorkbookCells(await readFile(outputPath)),
    ]);
    const patchLogPath = `${outputPath}.patch.json`;
    const log = {
        version: 1,
        createdAt: new Date().toISOString(),
        sourcePath: path,
        patches: diffCellMaps(before, after).map((entry) => ({
            id: entry.id,
            kind: 'formula',
            oldValue: entry.oldValue ?? '',
            newValue: entry.newValue ?? '',
        })),
    };
    await writePatchLog(patchLogPath, log);
    const validation = validate(after);
    return { ...result, outputPath, patchLog: patchLogPath, validation };
}
