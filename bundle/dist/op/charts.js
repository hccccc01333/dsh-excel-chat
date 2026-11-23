import { columnToNumber, numberToColumn, parseCellId } from '../formula.js';
import { plainContent } from '../workbook.js';
import { cellContentOf, parseNumericText, parseRange } from './core.js';
/** parseRange keeps the corners as written, so A5:A2 would otherwise be empty. */
function bounds(parsed) {
    return {
        top: Math.min(parsed.startRow, parsed.endRow),
        bottom: Math.max(parsed.startRow, parsed.endRow),
        left: Math.min(parsed.startCol, parsed.endCol),
        right: Math.max(parsed.startCol, parsed.endCol),
    };
}
function absoluteRef(parsed) {
    const sheet = parsed.sheet.name.replace(/'/g, "''");
    const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet}'`;
    const { top, bottom, left, right } = bounds(parsed);
    return `${quoted}!$${numberToColumn(left)}$${top}:$${numberToColumn(right)}$${bottom}`;
}
function linearCells(parsed) {
    const { top, bottom, left, right } = bounds(parsed);
    const rows = bottom - top + 1;
    const cols = right - left + 1;
    if (rows > 1 && cols > 1) {
        throw new Error('addChart ranges must be a single row or a single column');
    }
    return Array.from({ length: rows * cols }, (_, index) => parsed.sheet.getCell(top + (cols === 1 ? index : 0), left + (rows === 1 ? index : 0)));
}
function readSeries(parsed, nameOverride, fallbackName = '系列') {
    const cells = linearCells(parsed);
    const numbers = cells.map((cell) => {
        const numeric = typeof cell.value === 'number'
            ? cell.value
            : parseNumericText(plainContent(cellContentOf(cell)));
        if (numeric === null || !Number.isFinite(numeric)) {
            throw new Error(`addChart value ${cell.address} is not numeric`);
        }
        return numeric;
    });
    const { top, left } = bounds(parsed);
    const headerRow = top - 1;
    const header = headerRow >= 1
        ? plainContent(cellContentOf(parsed.sheet.getCell(headerRow, left)))
        : '';
    const firstValue = plainContent(cellContentOf(cells[0]));
    return {
        name: nameOverride || (header && header !== firstValue ? header : fallbackName),
        ref: absoluteRef(parsed),
        values: numbers,
    };
}
export function buildChartSpec(workbook, operation) {
    const categories = parseRange(workbook, operation.categories);
    const primary = parseRange(workbook, operation.values);
    if (categories.sheet.name !== primary.sheet.name) {
        throw new Error('addChart categories and values must be on the same sheet');
    }
    const categoryCells = linearCells(categories);
    const labels = categoryCells.map((cell) => plainContent(cellContentOf(cell)));
    const series = [readSeries(primary, undefined, '系列 1')];
    let rightmost = bounds(primary).right;
    for (const [index, extra] of (operation.series ?? []).entries()) {
        const parsed = parseRange(workbook, extra.values);
        if (parsed.sheet.name !== primary.sheet.name) {
            throw new Error('addChart all series must be on the same sheet as the categories');
        }
        series.push(readSeries(parsed, extra.name, `系列 ${index + 2}`));
        rightmost = Math.max(rightmost, bounds(parsed).right);
    }
    for (const item of series) {
        if (item.values.length !== labels.length) {
            throw new Error(`addChart ranges have different lengths: ${labels.length} vs ${item.values.length}`);
        }
    }
    const defaultAnchor = `${numberToColumn(Math.min(rightmost + 2, 16384))}${bounds(primary).top}`;
    const anchor = parseCellId(operation.anchor ?? `${primary.sheet.name}!${defaultAnchor}`);
    if (anchor.sheet.toUpperCase() !== primary.sheet.name.toUpperCase()) {
        throw new Error(`addChart anchor must be on the values sheet ${primary.sheet.name}`);
    }
    const anchorCol = columnToNumber(anchor.column);
    if (anchorCol < 1 || anchorCol > 16384 || anchor.row < 1 || anchor.row > 1048576) {
        throw new Error(`addChart anchor ${operation.anchor} is outside the worksheet`);
    }
    const widthPx = operation.width ?? 480;
    const heightPx = operation.height ?? 300;
    if (widthPx <= 0 || heightPx <= 0) {
        throw new Error('addChart width and height must be positive');
    }
    return {
        type: operation.type ?? 'column',
        title: operation.title,
        categoryRef: absoluteRef(categories),
        categories: labels,
        series,
        legend: operation.legend,
        xAxisTitle: operation.xAxisTitle,
        yAxisTitle: operation.yAxisTitle,
        anchorCol: anchorCol - 1,
        anchorRow: anchor.row - 1,
        width: Math.round(widthPx * 9525),
        height: Math.round(heightPx * 9525),
    };
}
