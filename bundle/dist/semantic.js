import { readFile } from 'node:fs/promises';
import { profileWorkbook } from './profile.js';
import { readWorkbookCells } from './workbook.js';
const TIME_PATTERN = /日期|时间|年月|季度|月份|date|time|month|year/i;
const ID_PATTERN = /编号|单号|序号|订单号|流水|代码|id|code/i;
const MEASURE_PATTERN = /金额|数量|销售额|成本|利润|毛利|单价|合计|总计|收入|费用|库存|价格|revenue|cost|profit|amount|qty|price|sales|total|sum/i;
function columnRole(profile) {
    const header = profile.header ?? '';
    if (profile.dtype === 'date' || TIME_PATTERN.test(header))
        return 'time';
    if (ID_PATTERN.test(header))
        return 'id';
    if (MEASURE_PATTERN.test(header) || profile.dtype === 'number')
        return 'measure';
    if (profile.dtype === 'string' && header !== '')
        return 'dimension';
    return 'unknown';
}
/**
 * Workbook semantic layer (v0.36 first slice): classify each column's role
 * (time / measure / dimension / id), derive sheet grain and formula-based
 * metrics, and find cross-sheet join keys. The summary feeds the planner and
 * verifier so the agent stops guessing "is 地区 column B?".
 */
export async function buildWorkbookSemanticProfile(path, sheet) {
    const profile = await profileWorkbook(path, sheet);
    const cells = await readWorkbookCells(await readFile(path));
    const formulasBySheet = new Map();
    for (const [id, content] of Object.entries(cells)) {
        if (!content.startsWith('='))
            continue;
        const bang = id.lastIndexOf('!');
        const sheetName = bang >= 0 ? id.slice(0, bang) : 'Sheet1';
        const list = formulasBySheet.get(sheetName) ?? [];
        list.push(`${id.slice(bang + 1)} = ${content}`);
        formulasBySheet.set(sheetName, list);
    }
    const sheets = profile.sheets.map((sheetProfile) => {
        const columns = sheetProfile.columns.map((column) => ({
            column: column.column,
            header: column.header,
            role: columnRole(column),
            dtype: column.dtype,
        }));
        const dimensions = columns.filter((column) => column.role === 'dimension' || column.role === 'time');
        const grain = dimensions.slice(0, 3).map((column) => column.header ?? column.column).join(' × ') || '未识别';
        const headerByColumn = new Map(sheetProfile.columns.map((column) => [column.column.toUpperCase(), column.header]));
        const derivedMetrics = (formulasBySheet.get(sheetProfile.sheet) ?? []).map((entry) => {
            const match = /^([A-Za-z]{1,3})(\d+) = =(.*)$/.exec(entry);
            if (!match)
                return entry;
            const header = headerByColumn.get(match[1].toUpperCase()) ?? match[1];
            return `${header} = ${match[3]}`;
        }).slice(0, 5);
        return { sheet: sheetProfile.sheet, columns, grain, derivedMetrics };
    });
    const joinKeys = [];
    const seen = new Set();
    for (let i = 0; i < sheets.length; i++) {
        for (let j = i + 1; j < sheets.length; j++) {
            const left = sheets[i];
            const right = sheets[j];
            for (const leftColumn of left.columns) {
                const header = leftColumn.header;
                if (!header || header.length < 2)
                    continue;
                if (right.columns.some((column) => column.header === header)) {
                    const signature = `${left.sheet}|${right.sheet}|${header}`;
                    if (!seen.has(signature)) {
                        seen.add(signature);
                        joinKeys.push({ left: left.sheet, right: right.sheet, key: header });
                    }
                }
            }
        }
    }
    const summary = sheets.map((entry) => {
        const parts = [
            `${entry.sheet}：粒度=${entry.grain}`,
            `时间=${entry.columns.filter((column) => column.role === 'time').map((column) => column.header ?? column.column).join('/') || '无'}`,
            `维度=${entry.columns.filter((column) => column.role === 'dimension').map((column) => column.header ?? column.column).join('/') || '无'}`,
            `指标=${entry.columns.filter((column) => column.role === 'measure').map((column) => column.header ?? column.column).join('/') || '无'}`,
            `标识=${entry.columns.filter((column) => column.role === 'id').map((column) => column.header ?? column.column).join('/') || '无'}`,
        ];
        if (entry.derivedMetrics.length > 0)
            parts.push(`派生=${entry.derivedMetrics.slice(0, 3).join('；')}`);
        return parts.join('；');
    }).join('\n');
    const joinSummary = joinKeys.length > 0
        ? `可关联：${joinKeys.map((key) => `${key.left}.${key.key} ↔ ${key.right}.${key.key}`).join('、')}`
        : '未发现跨表关联键';
    return { sheets, joinKeys, summary: `${summary}\n${joinSummary}` };
}
