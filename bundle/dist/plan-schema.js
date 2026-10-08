/**
 * A plan that cannot be salvaged. Carries the failure kind so the benchmark's
 * taxonomy can classify it without pattern-matching the message text — that
 * coupling made the messages untranslatable, because translating one silently
 * turned a planning error into an execution error.
 */
export class PlanSchemaError extends Error {
    // Declared explicitly rather than as a constructor parameter property: Node
    // runs TypeScript in strip-only mode, which rejects that syntax at load time
    // (`ERR_UNSUPPORTED_TYPESCRIPT_SYNTAX`) even though `tsc` compiles it happily.
    kind;
    constructor(kind, message) {
        super(message);
        this.name = 'PlanSchemaError';
        this.kind = kind;
    }
}
import { t } from './i18n.js';
const CELL_OR_RANGE = /^[A-Za-z]{1,3}\d+$|^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$/;
/** Ops whose `name` field references an EXISTING sheet (new-name ops excluded). */
const EXISTING_SHEET_NAME_OPS = new Set(['deleteSheet', 'hideSheet', 'duplicateSheet', 'setTabColor', 'moveSheet']);
/**
 * Salvage colloquial sheet references ("订单表") onto the exact sheet list
 * ("订单") when one contains the other; leave unknown names untouched so
 * genuinely wrong names still fail loudly.
 */
function matchSheetName(value, sheetNames) {
    if (sheetNames.some((sheet) => sheet === value))
        return value;
    const hit = sheetNames.find((sheet) => value.includes(sheet) || sheet.includes(value));
    return hit ?? value;
}
const REQUIRED_STRINGS = {
    fill: ['source', 'target'],
    fillSeries: ['start', 'target'],
    copyRange: ['source', 'target'],
    transpose: ['source', 'target'],
    copyStyle: ['source', 'target'],
    freezeFormulas: ['range'],
    clearRange: ['range'],
    uniqueValues: ['source', 'target'],
    dedupeRows: ['sheet'],
    fillMissing: ['range'],
    removeEmptyRows: ['range'],
    removeEmptyColumns: ['range'],
    trimText: ['range'],
    changeCase: ['range'],
    normalizeText: ['range'],
    splitColumn: ['sheet', 'column'],
    sortRange: ['range'],
    filterToRange: ['source', 'target'],
    style: ['range'],
    merge: ['range'],
    unmerge: ['range'],
    subtotal: ['sheet', 'range', 'groupColumn'],
    aggregateReport: ['source', 'groupColumn'],
    report: ['source', 'groupColumn'],
    preset: ['source', 'groupColumn'],
    highlightRows: ['sheet', 'range'],
    fuzzyMatch: ['source', 'target', 'valueColumn', 'outputColumn'],
    findReplace: ['find', 'replace'],
    addSheet: ['name'],
    deleteSheet: ['name'],
    renameSheet: ['oldName', 'newName'],
    duplicateSheet: ['name', 'newName'],
    mailMerge: ['template', 'data'],
    hideRows: ['sheet'],
    hideColumns: ['sheet'],
    groupRows: ['sheet'],
    groupColumns: ['sheet', 'from', 'to'],
    autoFitColumnWidths: ['sheet'],
    unfreezePanes: ['sheet'],
    unmergeAll: ['sheet'],
    setZoom: ['sheet'],
    showGridLines: ['sheet'],
    headerFooter: ['sheet'],
    printTitles: ['sheet'],
    moveSheet: ['name'],
    setHyperlink: ['cell'],
    joinSheets: ['source', 'sourceKey', 'lookup', 'lookupKey'],
    rankColumn: ['range', 'metricColumn', 'outputColumn'],
    rowPageBreaks: ['sheet'],
    clearPageBreaks: ['sheet'],
    addComment: ['cell', 'text'],
    addSparklines: ['dataRange', 'locationRange'],
    insertImage: ['cell'],
};
const REQUIRED_ARRAYS = {
    sortRange: ['keys'],
    filterToRange: ['criteria'],
    aggregateReport: ['metrics'],
    report: ['metrics'],
    preset: ['metrics'],
    subtotal: ['summaryColumns'],
    highlightRows: ['criteria'],
    conditionalFormatting: ['rules'],
    joinSheets: ['valueColumns', 'outputColumns'],
    hideColumns: ['columns'],
    rowPageBreaks: ['rows'],
};
const REQUIRED_NUMBERS = {
    freezePanes: ['row'],
    splitColumn: ['startRow'],
    setColumnWidth: ['width'],
    setRowHeight: ['row', 'height'],
    hideRows: ['from', 'to'],
    groupRows: ['start', 'end'],
    setZoom: ['zoom'],
    moveSheet: ['position'],
};
const REQUIRED_STRING_EXTRA = {
    freezePanes: ['sheet', 'column'],
};
/**
 * Normalize planner-produced assertions into WorkbookAssertions (Verifier 2.0):
 * keep only fully-formed entries, coerce ids to strings, and drop style-only
 * assertions that carry no checkable property. Invalid-but-salvageable fields
 * are repaired; entries with neither an id nor any expected value are dropped.
 */
export function sanitizeAssertions(assertions, sheetNames) {
    const notes = [];
    if (assertions === undefined || assertions === null)
        return { assertions: [], notes };
    if (!Array.isArray(assertions)) {
        notes.push(t('断言不是数组，已丢弃'));
        return { assertions: [], notes };
    }
    const out = [];
    assertions.forEach((entry, index) => {
        if (entry === null || typeof entry !== 'object' || Array.isArray(entry)) {
            notes.push(t('断言[{index}] 不是对象，已丢弃', { index }));
            return;
        }
        const raw = { ...entry };
        if (typeof raw.id !== 'string' || raw.id.trim() === '') {
            notes.push(t('断言[{index}] 缺少 id，已丢弃', { index }));
            return;
        }
        const id = raw.id.trim();
        // Bare cell/range ids get the first sheet prefix, matching op salvage.
        const body = id.includes('!') ? id.slice(id.lastIndexOf('!') + 1) : id;
        let normalizedId = id;
        if (!id.includes('!')) {
            if (CELL_OR_RANGE.test(id) || /^([A-Za-z]{1,3})(\d+)$/.test(id)) {
                normalizedId = `${sheetNames[0] ?? 'Sheet1'}!${id}`;
                notes.push(t('断言[{index}] 的 id 已补工作表前缀', { index }));
            }
        }
        const result = { id: normalizedId };
        let hasCheck = false;
        if (raw.expect !== undefined) {
            if (raw.expect === null) {
                result.expect = null;
                hasCheck = true;
            }
            else if (typeof raw.expect === 'string' || typeof raw.expect === 'number' || typeof raw.expect === 'boolean') {
                // Content checks compare against the serialized cell text.
                result.expect = String(raw.expect);
                hasCheck = true;
            }
            else {
                notes.push(t('断言[{index}] 的 expect 类型不支持，已忽略该字段', { index }));
            }
        }
        if (raw.startsWith !== undefined) {
            if (typeof raw.startsWith === 'string' && raw.startsWith !== '') {
                result.startsWith = raw.startsWith;
                hasCheck = true;
            }
            else {
                notes.push(t('断言[{index}] 的 startsWith 必须是非空字符串，已忽略', { index }));
            }
        }
        for (const key of ['fill', 'numberFormat', 'hAlign']) {
            const value = raw[key];
            if (typeof value === 'string' && value !== '') {
                result[key] = value;
                hasCheck = true;
            }
        }
        for (const key of ['bold', 'wrapText']) {
            const value = raw[key];
            if (typeof value === 'boolean') {
                result[key] = value;
                hasCheck = true;
            }
        }
        if (!hasCheck) {
            notes.push(t('断言[{index}] 没有可检查字段，已丢弃', { index }));
            return;
        }
        out.push(result);
    });
    return { assertions: out, notes };
}
/**
 * Validate and repair a planner-produced plan before execution. Salvageable
 * issues are fixed in place (sheet prefix, missing sheet, array wrapping,
 * alias fields, cell values); unsalvageable issues throw so the agent loop
 * can feed the exact message back to the planner for a corrected plan.
 */
export function sanitizePlan(steps, sheetNames) {
    const firstSheet = sheetNames[0] ?? 'Sheet1';
    const notes = [];
    const prefix = (value) => {
        if (typeof value !== 'string' || value.includes('!'))
            return value;
        return CELL_OR_RANGE.test(value) ? `${firstSheet}!${value}` : value;
    };
    const out = [];
    steps.forEach((step, stepIndex) => {
        if (!Array.isArray(step.operations)) {
            throw new PlanSchemaError('planning', t('第 {step} 步没有 operations 数组', { step: stepIndex + 1 }));
        }
        const operations = step.operations.map((operation, opIndex) => {
            if (!operation || typeof operation.op !== 'string') {
                throw new PlanSchemaError('planning', t('第 {step} 步第 {op} 个操作缺少 op 字段', { step: stepIndex + 1, op: opIndex + 1 }));
            }
            const raw = { ...operation };
            for (const key of ['range', 'source', 'target', 'start']) {
                const value = Array.isArray(raw[key]) ? raw[key][0] : raw[key];
                if (value !== undefined && typeof value === 'string' && !value.includes('!') && CELL_OR_RANGE.test(value)) {
                    raw[key] = `${firstSheet}!${value}`;
                    notes.push(t('{op} 的 {key} 已补工作表前缀', { op: operation.op, key }));
                }
            }
            if (raw.sheet === undefined && (REQUIRED_STRINGS[operation.op] ?? []).includes('sheet')) {
                raw.sheet = firstSheet;
                notes.push(t('{op} 已补默认工作表', { op: operation.op }));
            }
            // Salvage colloquial sheet references onto the exact sheet list.
            if (typeof raw.sheet === 'string') {
                const matched = matchSheetName(raw.sheet, sheetNames);
                if (matched !== raw.sheet) {
                    raw.sheet = matched;
                    notes.push(t('{op} 的 sheet 已匹配为 {matched}', { op: operation.op, matched }));
                }
            }
            if (EXISTING_SHEET_NAME_OPS.has(operation.op) && typeof raw.name === 'string') {
                const matched = matchSheetName(raw.name, sheetNames);
                if (matched !== raw.name) {
                    raw.name = matched;
                    notes.push(t('{op} 的 name 已匹配为 {matched}', { op: operation.op, matched }));
                }
            }
            if (operation.op === 'renameSheet' && typeof raw.oldName === 'string') {
                const matched = matchSheetName(raw.oldName, sheetNames);
                if (matched !== raw.oldName) {
                    raw.oldName = matched;
                    notes.push(t('renameSheet 的 oldName 已匹配为 {matched}', { matched }));
                }
            }
            for (const key of ['sheet', 'column', 'groupColumn', 'valueColumn', 'outputColumn', 'sourceKey', 'targetKey', 'name', 'oldName', 'newName', 'template', 'data', 'find', 'replace', 'delimiter']) {
                const value = raw[key];
                if (value === undefined)
                    continue;
                const scalar = Array.isArray(value) ? value[0] : value;
                if (scalar !== undefined && typeof scalar !== 'string') {
                    raw[key] = String(scalar);
                    notes.push(t('{op} 的 {key} 已转为字符串', { op: operation.op, key }));
                }
            }
            for (const field of ['metrics', 'summaryColumns', 'keys', 'criteria']) {
                const list = raw[field];
                if (!Array.isArray(list))
                    continue;
                raw[field] = list.map((entry, index) => {
                    if (entry === null || typeof entry !== 'object') {
                        throw new PlanSchemaError('planning', t('{op} 的 {field}[{index}] 必须是对象', { op: operation.op, field, index }));
                    }
                    const record = { ...entry };
                    if (record.column !== undefined && typeof record.column !== 'string') {
                        record.column = String(record.column);
                        notes.push(t('{op} 的 {field}[{index}].column 已转为字符串', { op: operation.op, field, index }));
                    }
                    if (record.column === undefined || record.column === '') {
                        throw new PlanSchemaError('planning', t('{op} 的 {field}[{index}].column 缺失', { op: operation.op, field, index }));
                    }
                    return record;
                });
            }
            if (operation.op === 'freezePanes' && raw.row === undefined && typeof raw.range === 'string') {
                const body = raw.range.includes('!') ? raw.range.slice(raw.range.lastIndexOf('!') + 1) : raw.range;
                const match = /^([A-Za-z]{1,3})(\d+)$/.exec(body);
                if (match) {
                    raw.column = match[1];
                    raw.row = Number(match[2]);
                    notes.push(t('freezePanes 已从 range {range} 推导 row/column', { range: String(raw.range) }));
                }
            }
            if (operation.op === 'crosstab') {
                // Planners sometimes emit flat metric fields instead of the object.
                if (raw.metric === undefined && (raw.metricColumn !== undefined || raw.metricFunction !== undefined)) {
                    raw.metric = { column: raw.metricColumn, function: raw.metricFunction ?? 'sum' };
                    notes.push(t('crosstab 的 metricColumn/metricFunction 已合并为 metric 对象'));
                }
                if (raw.metric !== undefined && raw.metric !== null && typeof raw.metric === 'object' && !Array.isArray(raw.metric)) {
                    const metric = { ...raw.metric };
                    if (metric.function === undefined) {
                        metric.function = 'sum';
                        notes.push(t('crosstab 的 metric.function 已补默认 sum'));
                    }
                    if (typeof metric.function !== 'string' || !['sum', 'average', 'count', 'counta', 'max', 'min'].includes(metric.function)) {
                        throw new PlanSchemaError('argument', t('crosstab 的 metric.function 不支持：{value}', { value: String(metric.function) }));
                    }
                    if (metric.column !== undefined && typeof metric.column !== 'string') {
                        metric.column = String(metric.column);
                        notes.push(t('crosstab 的 metric.column 已转为字符串'));
                    }
                    raw.metric = metric;
                }
            }
            if (operation.op === 'style' && raw.style !== undefined && raw.style !== null && typeof raw.style === 'object') {
                // Salvage exceljs-native alignment names onto the DSL fields.
                const style = { ...raw.style };
                if (style.hAlign === undefined && style.horizontal !== undefined) {
                    style.hAlign = style.horizontal;
                    notes.push(t('style 的 horizontal 已改名为 hAlign'));
                }
                if (style.vAlign === undefined && style.vertical !== undefined) {
                    style.vAlign = style.vertical;
                    notes.push(t('style 的 vertical 已改名为 vAlign'));
                }
                raw.style = style;
            }
            for (const key of REQUIRED_NUMBERS[operation.op] ?? []) {
                const value = raw[key];
                if (value === undefined || value === null || value === '') {
                    throw new PlanSchemaError('planning', t('{op} 缺少必填数字 {key}', { op: operation.op, key }));
                }
                const number = typeof value === 'string' ? Number(value) : value;
                if (typeof number !== 'number' || !Number.isFinite(number)) {
                    throw new PlanSchemaError('argument', t('{op} 的 {key} 必须是数字', { op: operation.op, key }));
                }
                raw[key] = number;
            }
            if (operation.op === 'fill' || operation.op === 'fillSeries') {
                const start = raw.start ?? raw.source;
                const target = raw.target;
                const targetBody = typeof target === 'string'
                    ? (target.includes('!') ? target.slice(target.lastIndexOf('!') + 1) : target)
                    : '';
                if (typeof start === 'string' && typeof target === 'string' && !targetBody.includes(':') && CELL_OR_RANGE.test(targetBody)) {
                    raw.target = `${start}:${targetBody}`;
                    notes.push(t('{op} 的 target 已扩展为 {target}', { op: operation.op, target: `${start}:${targetBody}` }));
                }
            }
            for (const field of REQUIRED_ARRAYS[operation.op] ?? []) {
                if (raw[field] === undefined) {
                    throw new PlanSchemaError('planning', t('{op} 缺少必填数组 {field}', { op: operation.op, field }));
                }
                if (!Array.isArray(raw[field])) {
                    raw[field] = [raw[field]];
                    notes.push(t('{op} 的 {field} 已包装为数组', { op: operation.op, field }));
                }
            }
            for (const field of REQUIRED_STRINGS[operation.op] ?? []) {
                if (raw[field] === undefined || raw[field] === null || raw[field] === '') {
                    throw new PlanSchemaError('planning', t('{op} 缺少必填字段 {field}', { op: operation.op, field }));
                }
            }
            for (const field of REQUIRED_STRING_EXTRA[operation.op] ?? []) {
                if (raw[field] === undefined || raw[field] === null || raw[field] === '') {
                    throw new PlanSchemaError('planning', t('{op} 缺少必填字段 {field}', { op: operation.op, field }));
                }
            }
            if (raw.cells && typeof raw.cells === 'object') {
                if (Array.isArray(raw.cells)) {
                    raw.cells = raw.cells.map((id) => prefix(id));
                }
                else if (operation.op === 'set') {
                    raw.cells = Object.fromEntries(Object.entries(raw.cells).map(([id, content]) => [prefix(id), String(content)]));
                }
            }
            if (operation.op === 'fillMissing') {
                if (raw.mode === undefined) {
                    raw.mode = 'value';
                    notes.push(t('fillMissing 已补 mode=value'));
                }
                if (raw.value === undefined && raw.fillValue !== undefined) {
                    raw.value = raw.fillValue;
                    notes.push(t('fillMissing 的 fillValue 已改为 value'));
                }
            }
            if (operation.op === 'renameSheet') {
                if (raw.oldName === undefined && raw.sheet !== undefined)
                    raw.oldName = raw.sheet;
                if (raw.newName === undefined && raw.target !== undefined)
                    raw.newName = raw.target;
            }
            if (['addSheet', 'deleteSheet', 'hideSheet', 'setTabColor', 'protectSheet', 'unprotectSheet', 'duplicateSheet'].includes(operation.op)) {
                if (raw.name === undefined && raw.sheet !== undefined)
                    raw.name = raw.sheet;
                if (operation.op === 'duplicateSheet' && raw.newName === undefined && raw.target !== undefined)
                    raw.newName = raw.target;
            }
            if (operation.op === 'filterToRange' && typeof raw.target === 'string' && !raw.target.includes('!') && !CELL_OR_RANGE.test(raw.target)) {
                raw.target = `${raw.target}!A1`;
                notes.push(t('filterToRange 的 target 已补 !A1'));
            }
            return raw;
        });
        out.push({ name: step.name, operations });
    });
    return { steps: out, notes };
}
/**
 * Run the same validation and salvage the planner path uses, for an operations
 * array that arrived straight from a tool call rather than from the planner.
 *
 * `excel_operate` and `excel_workflow` hand the model's array to the executor
 * directly, and the executor assumes it is already well-formed. So a model that
 * omitted a nested field reached a handler and crashed with
 * `TypeError: Cannot read properties of undefined (reading 'toUpperCase')`,
 * where the planner gets the actionable `sortRange 的 keys[0].column 缺失`.
 * Reusing this function rather than writing a second validator is deliberate:
 * one rule with two implementations drifts.
 */
export function sanitizeOperations(operations, sheetNames) {
    const [step] = sanitizePlan([{ operations }], sheetNames).steps;
    return step?.operations ?? [];
}
