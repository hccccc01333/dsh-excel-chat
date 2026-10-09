import ExcelJS from 'exceljs';
import { columnToNumber, numberToColumn, parseCellId, } from './formula.js';
import { listJoin, t } from './i18n.js';
import { applyPreserved, capturePreserved, isEmpty } from './preserve.js';
import { validateStructure } from './package-check.js';
import { findUnpreservedParts, writeWorkbookSafely } from './safe-write.js';
import { validate } from './validator.js';
import { readWorkbookCells, stripPivotTableParts } from './workbook.js';
import { diffCellMaps, writePatchLog } from './diff.js';
import { annotateWorkbookXml, emptyAnnotations } from './xml-postprocess.js';
import { readFile } from 'node:fs/promises';
export * from './operation-types.js';
import { findSheet, resolveCell, writeContent, parseRange, cellContentOf, qualifySheetName, properCase, normalizeTextValue, similarity, splitByWidth } from './op/core.js';
import { deleteColumnsFromSheet, deleteRowsFromSheet, shiftWorkbookColumns, shiftWorkbookRows } from './op/structure.js';
import { sortRange } from './op/sorting.js';
import { addTable, applyConditionalFormatting, applyDataValidation, applyStyle, normalizeColor } from './op/formatting.js';
import { applyFilterToRange, applyMailMerge, matchesCriterion } from './op/filters.js';
import { exportCsv, importCsv } from './op/csv.js';
import { insertImageIntoSheet } from './op/images.js';
import { applyFill, applyMerge, clearRange, copyRange, duplicateSheet, fillSeries, findReplace, renameSheetReferences, transposeRange } from './op/editing.js';
import { applyPageSetup, displayTextOf, displayWidth } from './op/layout.js';
import { joinSheets } from './op/joins.js';
import { applyCrosstab } from './op/crosstab.js';
import { applyRankColumn, applySheetView, copyStyle, freezeFormulas, moveSheet, setHyperlink, uniqueValues } from './op/sheets.js';
import { applyAggregateReport, applyPreset, applyReport, applySubtotal } from './op/reports.js';
// These three were part of this module's public surface before the split. They are
// re-exported rather than removed so every existing caller keeps importing them
// from here, and the split stays a change of layout rather than of interface.
export { findSheet, qualifySheetName, shiftFormulaReferences } from './op/core.js';
export async function applyOperationsToWorkbook(inputPath, operations, outputPath) {
    const original = await readFile(inputPath);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(stripPivotTableParts(original));
    const warnings = [];
    const annotations = emptyAnnotations();
    // Pivot tables are carried across the rewrite (see preserve.ts). Everything
    // else ExcelJS cannot round-trip is reported, because it really is lost.
    // `op: -1` marks a warning about the workbook itself rather than about one
    // operation in the list.
    const preserved = capturePreserved(original);
    const unpreserved = findUnpreservedParts(original);
    if (unpreserved.length > 0) {
        warnings.push({
            op: -1,
            message: t('这个文件含有本插件无法保留的内容，编辑后会丢失：{features}。需要保留请先另存一份副本，或改用原生 Excel 操作。', {
                features: listJoin(unpreserved),
            }),
        });
    }
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
                if (operation.widths && operation.delimiter !== undefined) {
                    throw new Error('splitColumn takes either delimiter or widths, not both');
                }
                if (!operation.widths && operation.delimiter === undefined) {
                    throw new Error('splitColumn requires delimiter or widths');
                }
                const columnNumber = columnToNumber(operation.column);
                const endRow = operation.endRow ?? sheet.rowCount;
                const partsByRow = new Map();
                let maxParts = 1;
                for (let row = operation.startRow; row <= endRow; row++) {
                    const text = cellContentOf(sheet.getCell(`${operation.column}${row}`));
                    if (!text)
                        continue;
                    const parts = operation.widths
                        ? splitByWidth(text, operation.widths)
                        : text.split(operation.delimiter).map((part) => part.trim());
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
                const mode = operation.widths
                    ? `fixed width ${operation.widths.join('/')}`
                    : `delimiter "${operation.delimiter}"`;
                warnings.push({ op: index, message: `splitColumn split ${partsByRow.size} row(s) into up to ${maxParts} columns (${mode})` });
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
                // Materializing a hidden empty row needs a height, so clamping to the
                // used range keeps huge "hide to the bottom" requests from bloating
                // the file with thousands of synthetic rows.
                const to = hidden ? Math.min(operation.to, sheet.rowCount) : operation.to;
                if (hidden && to < operation.to) {
                    warnings.push({ op: index, message: `hideRows clamped to the used range (row ${to})` });
                }
                for (let row = operation.from; row <= to; row++) {
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
                const level = operation.level ?? 1;
                if (operation.start < 1 || operation.end < operation.start)
                    throw new Error(`invalid groupRows: start=${operation.start} end=${operation.end}`);
                // Same clamp as hideRows: collapsed groups materialize empty rows.
                const end = level > 0 ? Math.min(operation.end, Math.max(sheet.rowCount, operation.start)) : operation.end;
                if (level > 0 && end < operation.end) {
                    warnings.push({ op: index, message: `groupRows clamped to the used range (row ${end})` });
                }
                for (let row = operation.start; row <= end; row++) {
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
                const { frozen, skipped } = freezeFormulas(workbook, operation.range);
                warnings.push({
                    op: index,
                    message: skipped > 0
                        ? t('freezeFormulas 转换 {frozen} 个公式，跳过 {skipped} 个无缓存结果的（先在 Excel 中打开计算后可再转）', { frozen, skipped })
                        : t('freezeFormulas 转换 {frozen} 个公式为缓存值', { frozen }),
                });
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
            case 'showFormulas': {
                const sheet = findSheet(workbook, operation.sheet);
                if (!sheet)
                    throw new Error(`sheet not found: ${operation.sheet}`);
                // Recorded as an annotation rather than set on the view: exceljs drops
                // showFormulas when it renders <sheetView>, so xml-postprocess injects it.
                if (operation.show === false)
                    annotations.showFormulas.delete(sheet.name);
                else
                    annotations.showFormulas.add(sheet.name);
                warnings.push({
                    op: index,
                    message: operation.show === false
                        ? `showFormulas off for ${sheet.name}`
                        : `showFormulas on for ${sheet.name} (the file opens showing formulas)`,
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
                const locBang = operation.locationRange.lastIndexOf('!');
                if (dataBang < 0 || locBang < 0)
                    throw new Error('addSparklines dataRange and locationRange must be sheet-qualified');
                const sparkSheet = findSheet(workbook, operation.dataRange.slice(0, dataBang));
                if (!sparkSheet)
                    throw new Error(`sheet not found: ${operation.dataRange.slice(0, dataBang)}`);
                const qualified = (range, bang) => `${qualifySheetName(sparkSheet.name)}!${range.slice(bang + 1)}`;
                if (operation.locationRange.slice(0, locBang).replace(/^'|'$/g, '') !== sparkSheet.name) {
                    throw new Error('sparkline dataRange and locationRange must be on the same sheet');
                }
                const groups = annotations.sparklines.get(sparkSheet.name) ?? [];
                groups.push({
                    dataRange: qualified(operation.dataRange, dataBang),
                    locationRange: qualified(operation.locationRange, locBang),
                    type: operation.type ?? 'line',
                    color: normalizeColor(operation.color ?? '375623'),
                    negativeColor: normalizeColor(operation.negativeColor ?? 'D00000'),
                    markers: operation.markers ?? false,
                    highColor: normalizeColor(operation.highColor ?? 'FF7C00'),
                    lowColor: normalizeColor(operation.lowColor ?? 'D00000'),
                });
                annotations.sparklines.set(sparkSheet.name, groups);
                break;
            }
            case 'insertImage': {
                await insertImageIntoSheet(workbook, operation);
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
                copyRange(workbook, operation.source, operation.target, operation.move ?? false, operation.valuesOnly ?? false, warnings, index);
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
    const sheetFileOf = new Map();
    if (annotations.comments.size > 0 || annotations.sparklines.size > 0 || annotations.showFormulas.size > 0) {
        workbook.eachSheet((sheet) => {
            sheetFileOf.set(sheet.name, `xl/worksheets/sheet${sheet.id}.xml`);
        });
    }
    const written = annotations.comments.size > 0 || annotations.sparklines.size > 0 || annotations.showFormulas.size > 0
        // ExcelJS cannot write comments or sparklines; inject the XML parts now.
        ? annotateWorkbookXml(new Uint8Array(buffer), annotations, sheetFileOf)
        : new Uint8Array(buffer);
    const merged = applyPreserved(written, preserved);
    if (!isEmpty(preserved)) {
        // The merge edits relationships across several files, and a mistake there
        // produces a package Excel will not open — worse than the pivot table it was
        // saving. Check before the write, and refuse rather than corrupt: nothing has
        // touched the target yet, so refusing costs the user an error message.
        const problems = validateStructure(merged);
        if (problems.length > 0) {
            throw new Error(t('保留原有内容时生成的表格未通过完整性校验，已放弃写入以免损坏文件：{problems}', {
                problems: listJoin(problems.slice(0, 3), 'semicolon'),
            }));
        }
    }
    await writeWorkbookSafely(outputPath, merged);
    return { warnings };
}
/** Visible text of a cell for width estimation: formula cells use their cached result. */
/**
 * Approximate display width in character units: CJK/fullwidth characters count
 * as 2 columns, everything else as 1.
 */
/** Join keys are compared trimmed + lowercased, numbers via their text form. */
/** Aggregations where a grand total of the computed grid is meaningful. */
/**
 * Fixed-width split: take `widths[i]` characters per output column. Anything
 * past the last width becomes a trailing column rather than being dropped.
 */
export async function operateWorkbookFile(path, operations, outputPath) {
    // Read the "before" state first. Reading it after the write only works when the
    // output is a different file; for an in-place edit (`outputPath === path`) both
    // reads returned the edited workbook, so the audit log came out empty and the
    // patch-log rollback had nothing to replay.
    const before = await readWorkbookCells(await readFile(path));
    const result = await applyOperationsToWorkbook(path, operations, outputPath);
    const after = await readWorkbookCells(await readFile(outputPath));
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
