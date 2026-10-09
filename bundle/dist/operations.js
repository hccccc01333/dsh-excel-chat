import ExcelJS from 'exceljs';
import { columnToNumber, numberToColumn, parseCellId, } from './formula.js';
import { guardFormulaInjection, parseCsv, stringifyCsv, unguardFormulaInjection } from './csv.js';
import { listJoin, t } from './i18n.js';
import { applyPreserved, capturePreserved, isEmpty } from './preserve.js';
import { validateStructure } from './package-check.js';
import { findUnpreservedParts, writeWorkbookSafely } from './safe-write.js';
import { validate } from './validator.js';
import { cellContent, contentToCellValue, plainContent, readWorkbookCells, stripPivotTableParts } from './workbook.js';
import { diffCellMaps, writePatchLog } from './diff.js';
import { annotateWorkbookXml, emptyAnnotations } from './xml-postprocess.js';
import { readFile, writeFile } from 'node:fs/promises';
export * from './operation-types.js';
import { RANGE_LINE, findSheet, resolveCell, writeContent, parseRange, shiftFormulaReferences, cellContentOf, qualifySheetName, absoluteColumnRef, properCase, normalizeTextValue, similarity, parseTargetCell } from './op/core.js';
import { deleteColumnsFromSheet, deleteRowsFromSheet, shiftWorkbookColumns, shiftWorkbookRows } from './op/structure.js';
import { sortRange } from './op/sorting.js';
import { addTable, applyConditionalFormatting, applyDataValidation, applyStyle, normalizeColor } from './op/formatting.js';
import { applyFilterToRange, applyMailMerge, matchesCriterion } from './op/filters.js';
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
function copyRange(workbook, sourceRange, targetCell, move, valuesOnly = false, warnings, opIndex = 0) {
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
    // Snapshot the source block before writing: copying onto an overlapping
    // range (down/right) would otherwise read cells that earlier writes already
    // replaced, corrupting the result.
    const snapshot = [];
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const source = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            snapshot.push({ row, col, formula: source.formula, result: source.formula ? source.result : undefined, value: source.value, content: cellContentOf(source) });
        }
    }
    for (const { row, col, formula, result, value, content } of snapshot) {
        const destCol = targetCol + (col - parsed.startCol);
        const destRow = targetRow + (row - parsed.startRow);
        const dest = targetSheet.getCell(`${numberToColumn(destCol)}${destRow}`);
        if (valuesOnly) {
            // Paste-special: values only. Formulas contribute their last cached
            // result; empty cells clear the destination.
            if (formula) {
                if (result === undefined || result === null) {
                    // No cached value (common for freshly written formulas): fall back
                    // to copying the shifted formula so nothing is lost.
                    dest.value = {
                        formula: shiftFormulaReferences(content, parsed.sheet.name, null, {
                            rowDelta: destRow - row,
                            colDelta: destCol - col,
                        }).slice(1),
                    };
                    warnings?.push({ op: opIndex, message: t('copyRange valuesOnly：部分公式无缓存结果，已按公式复制') });
                }
                else {
                    dest.value = result;
                }
            }
            else {
                dest.value = value;
            }
            continue;
        }
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
            : contentToCellValue(content);
    }
    if (move) {
        for (let row = parsed.startRow; row <= parsed.endRow; row++) {
            for (let col = parsed.startCol; col <= parsed.endCol; col++) {
                // Clear the source only for cells that the destination did not also
                // write into (fully non-overlapping copies); overlapping in-place moves
                // keep the copied block intact.
                const cleared = targetCol + (col - parsed.startCol);
                const clearedRow = targetRow + (row - parsed.startRow);
                const outsideDest = cleared < parsed.startCol || cleared > parsed.endCol ||
                    clearedRow < parsed.startRow || clearedRow > parsed.endRow ||
                    targetSheet.name !== parsed.sheet.name;
                if (outsideDest)
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
            const cell = sheet.getCell(`${numberToColumn(colIndex + 1)}${rowIndex + 1}`);
            const { text, guarded } = unguardFormulaInjection(value);
            // A guarded field was text when it was written, so it has to stay text:
            // handing it back to writeContent would infer `=` and revive the formula
            // the guard exists to defuse.
            if (guarded)
                cell.value = text;
            else
                writeContent(cell, value);
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
                // Serialise through `cellContent`, not `String(raw)`. Dates, hyperlinks,
                // rich text and error cells are all objects, and `String()` on those gave
                // a locale-and-timezone-dependent date string or a literal
                // `[object Object]` — four shapes written into the CSV as garbage.
                const raw = cell.value;
                let text = plainContent(cellContent(cell) ?? '');
                if (guard && typeof raw === 'string')
                    text = guardFormulaInjection(text);
                row.push(text);
            }
        }
        rows.push(row);
    }
    await writeFile(options.file, stringifyCsv(rows, options.delimiter ?? ','), 'utf8');
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
                    : contentToCellValue(replaced);
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
    const newBare = `${newName}!`;
    // Bare references need token boundaries so renaming "A" does not corrupt
    // "AA!" (oldBare is a substring of "AA!"). Match only when the name is not
    // preceded by an identifier char / quote / $ and is followed by a cell ref.
    const bareRef = new RegExp(`(?<![A-Za-z0-9_$'])${escapeRegExp(oldName)}!(?=[A-Za-z$])`, 'g');
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (!cell.formula)
                    return;
                const formula = cell.formula
                    .replaceAll(oldQuoted, newQuoted)
                    .replace(bareRef, () => newBare);
                if (formula !== cell.formula)
                    cell.value = { formula };
            });
        });
    });
}
function escapeRegExp(text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
    // Snapshot the whole source block first: transposing onto the source (or an
    // overlapping area) must not read cells that earlier writes already replaced.
    const snapshot = [];
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const source = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            snapshot.push({ row, col, content: cellContentOf(source), raw: source.value });
        }
    }
    // Transposing onto the source footprint: clear the source block first so a
    // non-square source (e.g. 2x4 -> 4x2) does not leave stale cells behind.
    if (target.sheet.name === parsed.sheet.name) {
        const destLastRow = target.row + (parsed.endCol - parsed.startCol);
        const destLastCol = target.col + (parsed.endRow - parsed.startRow);
        const overlaps = target.row <= parsed.endRow && destLastRow >= parsed.startRow &&
            target.col <= parsed.endCol && destLastCol >= parsed.startCol;
        if (overlaps) {
            for (let row = parsed.startRow; row <= parsed.endRow; row++) {
                for (let col = parsed.startCol; col <= parsed.endCol; col++) {
                    parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value = null;
                }
            }
        }
    }
    for (const { row, col, content, raw } of snapshot) {
        // (row,col) maps to (targetRow + colOffset, targetCol + rowOffset).
        const destRow = target.row + (col - parsed.startCol);
        const destCol = target.col + (row - parsed.startRow);
        const dest = target.sheet.getCell(`${numberToColumn(destCol)}${destRow}`);
        if (!content)
            continue;
        dest.value = content.startsWith('=')
            ? contentToCellValue(shiftFormulaReferences(content, parsed.sheet.name, null, {
                rowDelta: destRow - row,
                colDelta: destCol - col,
            }))
            : raw;
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
                value.startsWith('=') ? { formula: value.slice(1) } : contentToCellValue(value);
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
    // Keep the raw cell value so the output header/label cells match the source
    // criteria (dates as serials, numbers as numbers); text is only for dedup.
    const collectKeys = (col) => {
        const keys = [];
        const seen = new Set();
        for (let row = firstData; row <= parsed.endRow; row++) {
            const raw = parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value;
            const text = raw === null || raw === undefined ? '' : String(raw);
            if (!seen.has(text)) {
                seen.add(text);
                keys.push({ raw, text });
            }
        }
        return keys;
    };
    const rowKeys = collectKeys(rowCol);
    const colKeys = collectKeys(colCol);
    if (!rowKeys.length || !colKeys.length)
        throw new Error('crosstab source has no data rows');
    const sheetRange = (col) => absoluteColumnRef(parsed.sheet.name, numberToColumn(col), firstData, parsed.endRow);
    const rowRange = sheetRange(rowCol);
    const colRange = sheetRange(colCol);
    // SUMIFS/AVERAGEIFS/MAXIFS/MINIFS take a leading sum/average range; the
    // *IFS count form takes only criteria pairs, so metric.range must be dropped
    // there or the argument list becomes invalid (odd count of range/criteria).
    const usesMetric = options.metric.function !== 'count' && options.metric.function !== 'counta';
    const metricRange = usesMetric ? sheetRange(columnToNumber(options.metric.column)) : null;
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
        cell.value = key.raw === undefined || key.raw === null ? '' : key.raw;
        cell.font = { bold: true };
    });
    rowKeys.forEach((rowKey, rowIndex) => {
        const outRow = 2 + rowIndex;
        output.getCell(`A${outRow}`).value = rowKey.raw === undefined || rowKey.raw === null ? '' : rowKey.raw;
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
/** Formats exceljs can embed, and the size ceiling that keeps memory sane. */
const IMAGE_EXTENSIONS = new Set(['png', 'jpeg', 'jpg', 'gif']);
const MAX_IMAGE_BYTES = 20 * 1024 * 1024;
/**
 * Embed an image anchored at a cell. Cross-platform: exceljs writes the media
 * part and the drawing XML itself, so no Excel installation is involved. Because
 * we mutate the loaded workbook (rather than rebuilding it), images that were
 * already in the file survive — verified by reading the media parts back.
 */
async function insertImageIntoSheet(workbook, options) {
    const parsed = parseCellId(options.cell);
    const sheet = findSheet(workbook, parsed.sheet);
    if (!sheet)
        throw new Error(`sheet not found: ${parsed.sheet}`);
    if (options.file && options.base64)
        throw new Error('insertImage takes either file or base64, not both');
    if (!options.file && !options.base64)
        throw new Error('insertImage requires file or base64');
    const { buffer, extension } = await readImageSource(options);
    if (buffer.byteLength > MAX_IMAGE_BYTES) {
        throw new Error(`insertImage refuses images over ${MAX_IMAGE_BYTES / 1024 / 1024}MB (got ${Math.round(buffer.byteLength / 1024 / 1024)}MB)`);
    }
    // exceljs ships an older Buffer typing that Node 22's generic Buffer no longer
    // satisfies; the bytes are identical, so narrow it at the boundary.
    const imageId = workbook.addImage({ buffer, extension });
    // exceljs requires `ext` both in its types and at render time — omitting it
    // throws while writing the drawing XML.
    const ext = resolveImageExtent(options, buffer);
    if (!ext) {
        throw new Error('insertImage could not read the image size; pass width and height explicitly');
    }
    sheet.addImage(imageId, {
        tl: { col: columnToNumber(parsed.column) - 1, row: parsed.row - 1 },
        ext,
    });
}
/** Resolve image bytes plus exceljs's extension token, from a path or base64. */
async function readImageSource(options) {
    if (options.file) {
        // Check the extension first so a typo'd format fails fast, before touching disk.
        const extension = toImageExtension(options.file);
        // readFile yields Buffer<ArrayBufferLike> while exceljs types want the narrower
        // Buffer. Identical bytes, so a cast beats copying up to 20MB for nothing.
        return { buffer: await readFile(options.file), extension };
    }
    const raw = options.base64;
    const dataUri = /^data:image\/([a-z0-9]+);base64,(.*)$/is.exec(raw);
    if (dataUri) {
        return {
            buffer: Buffer.from(dataUri[2].replace(/\s+/g, ''), 'base64'),
            extension: toImageExtension(dataUri[1]),
        };
    }
    // Bare base64 carries no format hint; PNG is the safe default (and the caller
    // can pass a data URI whenever the payload is actually jpeg/gif).
    return { buffer: Buffer.from(raw.replace(/\s+/g, ''), 'base64'), extension: 'png' };
}
function toImageExtension(hint) {
    const extension = (/\.([a-z0-9]+)$/i.exec(hint)?.[1] ?? hint).toLowerCase();
    if (!IMAGE_EXTENSIONS.has(extension)) {
        throw new Error(`insertImage supports png/jpeg/gif, got "${extension}"`);
    }
    return extension === 'jpg' ? 'jpeg' : extension;
}
/**
 * Intrinsic pixel size from the image header. PNG and GIF keep it at a fixed
 * offset; JPEG needs a walk to the first SOF segment. Returns null for anything
 * unrecognised, so the caller asks for an explicit size instead of guessing.
 */
/**
 * The rendered size for an embedded image.
 *
 * Both sides given → use them as-is. Neither → the image's own pixel size. Just
 * one → scale the other from the intrinsic size, so asking for `width: 300` on a
 * 100×200 image renders 300×600 rather than squashing it to 300×100. Returns
 * null only when the caller gave a single side and the header is unreadable.
 */
function resolveImageExtent(options, buffer) {
    const { width, height } = options;
    if (width !== undefined && height !== undefined)
        return { width, height };
    const intrinsic = readImageSize(buffer);
    if (width === undefined && height === undefined)
        return intrinsic;
    if (!intrinsic)
        return null;
    if (width !== undefined) {
        return { width, height: scaleDimension(width, intrinsic.height, intrinsic.width) };
    }
    return { width: scaleDimension(height, intrinsic.width, intrinsic.height), height: height };
}
/** Proportional counterpart to a given dimension, never rounding down to zero. */
function scaleDimension(given, numerator, denominator) {
    if (denominator <= 0)
        return given;
    return Math.max(1, Math.round((given * numerator) / denominator));
}
function readImageSize(buffer) {
    // PNG: IHDR chunk holds big-endian u32 width/height at offsets 16 and 20.
    if (buffer.length >= 24 && buffer.readUInt32BE(0) === 0x89504e47) {
        return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
    }
    // GIF: logical screen descriptor, little-endian u16 at offsets 6 and 8.
    if (buffer.length >= 10 && buffer.toString('latin1', 0, 3) === 'GIF') {
        return { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8) };
    }
    // JPEG: walk marker segments to the frame header (SOF0-SOF15, minus DHT/JPG/DAC).
    if (buffer.length >= 4 && buffer[0] === 0xff && buffer[1] === 0xd8) {
        let offset = 2;
        while (offset + 9 <= buffer.length) {
            if (buffer[offset] !== 0xff) {
                offset++;
                continue;
            }
            const marker = buffer[offset + 1];
            if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
                return { width: buffer.readUInt16BE(offset + 7), height: buffer.readUInt16BE(offset + 5) };
            }
            const segmentLength = buffer.readUInt16BE(offset + 2);
            if (segmentLength < 2)
                break;
            offset += 2 + segmentLength;
        }
    }
    return null;
}
/**
 * Fixed-width split: take `widths[i]` characters per output column. Anything
 * past the last width becomes a trailing column rather than being dropped.
 */
function splitByWidth(text, widths) {
    const parts = [];
    let offset = 0;
    for (const width of widths) {
        parts.push(text.slice(offset, offset + width).trim());
        offset += width;
    }
    if (offset < text.length)
        parts.push(text.slice(offset).trim());
    return parts;
}
function setHyperlink(workbook, options) {
    const cell = resolveCell(workbook, options.cell);
    const text = options.text;
    if (options.url) {
        cell.value = { text: text ?? options.url, hyperlink: options.url };
        return;
    }
    if (options.location) {
        // ExcelJS mishandles internal links (writes both a "#..." location AND an
        // External relationship), so use a HYPERLINK() formula, which Excel always
        // navigates correctly. The location's sheet name is quoted if needed.
        const loc = internalLinkLocation(options.location);
        const label = text ?? options.location.replace(/^#/, '');
        cell.value = { formula: `HYPERLINK("${loc}","${label.replaceAll('"', '""')}")` };
        return;
    }
    throw new Error('setHyperlink requires url (external) or location (internal, e.g. "Sheet2!A1")');
}
/** Turn "Sheet2!A1" / "#明细!B2" into the HYPERLINK target "#'Sheet 2'!A1". */
function internalLinkLocation(location) {
    const bare = location.startsWith('#') ? location.slice(1) : location;
    const bang = bare.lastIndexOf('!');
    if (bang < 0)
        return `#${bare}`; // a defined name
    const sheet = bare.slice(0, bang).replaceAll(/^'|'$/g, '');
    const ref = bare.slice(bang + 1);
    return `#${qualifySheetName(sheet)}!${ref}`;
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
    let skipped = 0;
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            if (!cell.formula)
                continue;
            // Plugin-written formulas often carry no cached result (never opened in
            // Excel). Freezing those to "null" would erase them, so leave them as
            // formulas and skip.
            const result = cell.result;
            if (result === undefined || result === null) {
                skipped++;
                continue;
            }
            cell.value = result;
            frozen++;
        }
    }
    return { frozen, skipped };
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
        const key = uniqueValueKey(cell, raw);
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
    return ordered.length;
}
/**
 * Dedup key that does not collapse distinct types: number 1, text "1" and
 * boolean TRUE get separate keys; a formula with no cached result keys on its
 * formula text instead of collapsing every such cell to "".
 */
function uniqueValueKey(cell, raw) {
    if (cell.formula && (raw === undefined || raw === null))
        return `=f:${cell.formula}`;
    if (raw === null || raw === undefined)
        return '∅';
    if (raw instanceof Date)
        return `date:${raw.getTime()}`;
    if (typeof raw === 'object')
        return `obj:${JSON.stringify(raw)}`;
    return `${typeof raw}:${String(raw)}`;
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
    const metricRange = absoluteColumnRef(parsed.sheet.name, numberToColumn(metricCol), firstData, parsed.endRow);
    for (let row = firstData; row <= parsed.endRow; row++) {
        parsed.sheet.getCell(`${options.outputColumn}${row}`).value = {
            formula: `RANK(${numberToColumn(metricCol)}${row},${metricRange},${options.descending === false ? 1 : 0})`,
        };
    }
}
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
