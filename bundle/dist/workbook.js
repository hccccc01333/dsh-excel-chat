import ExcelJS from 'exceljs';
import { readFile } from 'node:fs/promises';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { validate } from './validator.js';
export function cellContent(cell) {
    if (cell.formula)
        return `=${cell.formula}`;
    const value = cell.value;
    if (value === null || value === undefined)
        return null;
    if (typeof value === 'object') {
        if (value instanceof Date)
            return value.toISOString();
        const text = value.text;
        if (typeof text === 'string')
            return text;
        const richText = value.richText;
        if (Array.isArray(richText))
            return cell.text ?? null;
        return JSON.stringify(value);
    }
    return String(value);
}
/**
 * ExcelJS crashes when a worksheet's `<tableParts>` points at a pivot table
 * (it only understands regular tables). Strip those anchors before loading so
 * read/validate/operate keep working on files that contain pivot tables.
 * Pivot parts are dropped on rewrite — acceptable, since ExcelJS cannot
 * preserve them anyway.
 */
export function stripPivotTableParts(data) {
    const files = unzipSync(data);
    const affectedSheets = [];
    for (const name of Object.keys(files)) {
        if (!/^xl\/worksheets\/_rels\/.*\.rels$/.test(name))
            continue;
        if (strFromU8(files[name]).includes('pivotTable')) {
            affectedSheets.push(`xl/worksheets/${name.split('/').pop().replace(/\.rels$/, '')}`);
        }
    }
    if (affectedSheets.length === 0)
        return data;
    for (const sheetFile of affectedSheets) {
        const xml = strFromU8(files[sheetFile] ?? new Uint8Array(0));
        if (!xml.includes('tableParts'))
            continue;
        files[sheetFile] = strToU8(xml.replace(/<tableParts[^>]*>[\s\S]*?<\/tableParts>/g, ''));
    }
    return Buffer.from(zipSync(files));
}
export async function readWorkbookCells(data) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(stripPivotTableParts(data));
    const cells = {};
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                const content = cellContent(cell);
                if (content !== null)
                    cells[`${sheet.name}!${cell.address}`] = content;
            });
        });
    });
    return cells;
}
export async function validateWorkbookFile(path) {
    const data = await readFile(path);
    return validate(await readWorkbookCells(data));
}
