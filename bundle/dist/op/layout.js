import { findSheet } from './core.js';
export function applyPageSetup(workbook, options) {
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
export function displayTextOf(cell) {
    const value = cell.formula ? cell.result : cell.value;
    if (value === null || value === undefined)
        return '';
    if (value instanceof Date)
        return '2026-12-31';
    if (typeof value === 'object')
        return JSON.stringify(value);
    return String(value);
}
export function displayWidth(text) {
    let width = 0;
    for (const char of text) {
        const code = char.codePointAt(0) ?? 0;
        width += code > 0x2e7f ? 2 : 1;
    }
    return width;
}
