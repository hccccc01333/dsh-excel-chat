import { readFile, writeFile } from 'node:fs/promises';
import { numberToColumn } from '../formula.js';
import { guardFormulaInjection, parseCsv, stringifyCsv, unguardFormulaInjection } from '../csv.js';
import { cellContent, plainContent } from '../workbook.js';
import { findSheet, parseRange, writeContent } from './core.js';
export async function importCsv(workbook, options) {
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
export async function exportCsv(workbook, options) {
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
