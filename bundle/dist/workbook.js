import ExcelJS from 'exceljs';
import { createHash } from 'node:crypto';
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
/**
 * The inverse of `cellContent`: turn a serialized cell string back into a typed
 * value. Shared so that every writer — `set`, `importCsv`, and the patch-log
 * rollback behind `excel_undo` — infers the same types from the same strings.
 *
 * The rollback used to assign the raw string, which turned every restored
 * number, boolean and date into *text*: undoing an edit left `42` as the string
 * "42". The damage was invisible because `cellContent` renders both the same
 * way, so comparing the strings showed no difference at all.
 */
export function contentToCellValue(content) {
    if (content.startsWith('='))
        return { formula: content.slice(1) };
    if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(content))
        return Number(content);
    if (/^true$/i.test(content))
        return true;
    if (/^false$/i.test(content))
        return false;
    // Accept the whole ISO 8601 date-time form, including the fractional seconds
    // and `Z` that `cellContent` emits — the reader's own output has to be
    // writable again, or reading a date and writing it back silently degrades the
    // cell to text.
    const date = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?(Z|[+-]\d{2}:?\d{2})?$/.exec(content);
    if (date) {
        const [, year, month, day, hour, minute, second, millis, zone] = date;
        if (zone) {
            // A zone designator makes the string an absolute instant. Re-deriving the
            // components in local time would shift it by the offset.
            const instant = new Date(content.replace(' ', 'T'));
            if (!Number.isNaN(instant.getTime()))
                return instant;
        }
        // Without a zone the value is a wall clock, which is how exceljs reads and
        // writes dates, so the components go in as local time.
        return new Date(Number(year), Number(month) - 1, Number(day), Number(hour ?? 0), Number(minute ?? 0), Number(second ?? 0), Number(millis ?? 0));
    }
    return content;
}
/**
 * `cellContent` wraps an error cell as `{"error":"#REF!"}` so an error stays
 * distinguishable from text that happens to spell the same token. A consumer
 * that wants the plain representation — a CSV, which Excel writes as the bare
 * token — unwraps it here instead of re-deriving the shape.
 */
export function plainContent(content) {
    const serialised = /^\{"error":"([^"]+)"\}$/.exec(content);
    return serialised ? serialised[1] : content;
}
/**
 * Sheet names in workbook order. The plan salvage needs them to fill in a
 * missing `sheet`/range prefix, so anything applying model-supplied operations
 * has to know the sheets before it can validate them.
 */
export async function readWorkbookSheetNames(path) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(stripPivotTableParts(await readFile(path)));
    return workbook.worksheets.map((sheet) => sheet.name);
}
/**
 * Content hash of a workbook's parts, used to answer "did this step change
 * anything?".
 *
 * It hashes the serialized parts rather than a snapshot of the cell grid,
 * because most of what an operation can change lives *outside* the grid:
 * comments, frozen panes, column widths, autofilters, conditional formats, data
 * validations, tables, images, sparklines, defined names, tab colours. A
 * cell-value snapshot reported "no change" for every one of those, and since
 * the agent loop forces `achieved: false` when nothing changed, a step that had
 * in fact worked was judged a failure. Hashing the parts is complete by
 * construction — a new operation is covered the moment it writes anything —
 * whereas an enumerated snapshot has to be extended by hand and drifts.
 *
 * Both sides are re-serialized through exceljs first. It normalizes on write,
 * so a file Excel wrote and one exceljs wrote differ even when they are
 * semantically identical; without this the comparison would always say
 * "changed". exceljs output is byte-deterministic (verified: two passes over
 * one input are identical, and `dcterms:created`/`modified` are preserved
 * rather than re-stamped), so this normalization is safe.
 */
export async function workbookFingerprint(path) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(stripPivotTableParts(await readFile(path)));
    const parts = unzipSync(new Uint8Array(await workbook.xlsx.writeBuffer()));
    const hash = createHash('sha256');
    for (const name of Object.keys(parts).sort()) {
        if (name.endsWith('/'))
            continue;
        hash.update(name);
        hash.update(parts[name]);
    }
    return hash.digest('hex');
}
