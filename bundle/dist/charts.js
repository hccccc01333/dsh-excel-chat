import { readFile } from 'node:fs/promises';
import { posix } from 'node:path';
import { unzipSync } from 'fflate';
import { canonicalCellId, columnToNumber, normalizeSheet, numberToColumn } from './formula.js';
export async function readChartInfos(path) {
    const files = unzipSync(new Uint8Array(await readFile(path)));
    const decoder = new TextDecoder();
    const text = (name) => {
        const data = files[name];
        return data ? decoder.decode(data) : undefined;
    };
    const workbookXml = text('xl/workbook.xml');
    const workbookRels = text('xl/_rels/workbook.xml.rels');
    if (!workbookXml || !workbookRels)
        return [];
    const relTargets = new Map();
    for (const match of workbookRels.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
        relTargets.set(match[1], match[2]);
    }
    const sheets = [];
    for (const match of workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
        const target = relTargets.get(match[2]);
        if (target)
            sheets.push({ name: match[1], file: target.replace(/^\//, '') });
    }
    const infos = [];
    for (const sheet of sheets) {
        const sheetBase = sheet.file.split('/').pop().replace(/\.xml$/, '');
        const rels = text(`xl/worksheets/_rels/${sheetBase}.xml.rels`);
        if (!rels)
            continue;
        for (const match of rels.matchAll(/<Relationship[^>]*Type="[^"]*\/chart"[^>]*Target="([^"]+)"/g)) {
            const chartPath = posix.resolve('/xl/worksheets', match[1]).slice(1);
            const chartXml = text(chartPath);
            if (!chartXml)
                continue;
            infos.push({ sheetName: sheet.name, chartPath, ...parseChartXml(chartXml) });
        }
    }
    return infos;
}
export function parseChartXml(xml) {
    const typeMatch = xml.match(/<c:(barChart|lineChart|pieChart|scatterChart|areaChart|doughnutChart|radarChart|bubbleChart|surfaceChart|stockChart)/);
    const series = [];
    for (const match of xml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
        const body = match[1];
        series.push({
            name: extractRef(body, 'c:tx'),
            categories: extractRef(body, 'c:cat'),
            values: extractRef(body, 'c:val'),
        });
    }
    return { type: typeMatch ? typeMatch[1] : null, series };
}
function extractRef(body, tag) {
    const section = new RegExp(`<${tag}>([\\s\\S]*?)<\\/${tag}>`).exec(body)?.[1];
    if (!section)
        return undefined;
    return /<c:f>([^<]+)<\/c:f>/.exec(section)?.[1]?.trim();
}
export function parseRangeRef(ref) {
    const match = /^(?:'([^']+)'|([^!]+))!?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(ref.trim());
    if (!match)
        return null;
    const startColumn = match[3].toUpperCase();
    const endColumn = (match[5] ?? match[3]).toUpperCase();
    const startRow = Number(match[4]);
    const endRow = match[6] ? Number(match[6]) : startRow;
    return {
        sheet: normalizeSheet(match[1] ?? match[2] ?? null),
        startColumn,
        startRow,
        endColumn,
        endRow,
    };
}
export function expandRange(range) {
    const colMin = Math.min(columnToNumber(range.startColumn), columnToNumber(range.endColumn));
    const colMax = Math.max(columnToNumber(range.startColumn), columnToNumber(range.endColumn));
    const rowMin = Math.min(range.startRow, range.endRow);
    const rowMax = Math.max(range.startRow, range.endRow);
    const ids = [];
    for (let col = colMin; col <= colMax; col++) {
        for (let row = rowMin; row <= rowMax; row++) {
            ids.push(canonicalCellId(range.sheet, numberToColumn(col), row));
        }
    }
    return ids;
}
