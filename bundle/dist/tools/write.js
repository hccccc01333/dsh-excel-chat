import { defineTool } from '@deepseek-ai/dsh-tools';
import { t } from '../i18n.js';
import { createPivotTable } from '../pivot.js';
import { readPatchLog, rollbackPatchLog } from '../diff.js';
import { restoreSnapshot } from '../safe-write.js';
import { excelOperationSchema } from '../operation-schema.js';
import { sanitizeOperations } from '../plan-schema.js';
import { operateWorkbookFile } from '../operations.js';
import { readWorkbookSheetNames } from '../workbook.js';
export function registerWriteTools(ctx) {
    ctx.effect(() => ctx.tools.register(defineTool({
        name: 'excel_create_pivot',
        description: 'Create a native Excel pivot table (pivotCache + pivotTable) from a source range: choose one or more row fields, optional column fields and report filters, plus value fields (sum/count/average/max/min). The pivot renders in Excel and can be refreshed when source data changes.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            sheet: {
                type: 'string',
                required: true,
                description: 'Source sheet name.',
            },
            range: {
                type: 'string',
                required: true,
                description: 'Source data range including the header row, e.g. "订单!A1:F7".',
            },
            rows: {
                type: 'array',
                items: { type: 'string' },
                required: true,
                description: 'Row field column letters, e.g. ["B"] or ["B","C"].',
            },
            columns: {
                type: 'array',
                items: { type: 'string' },
                description: 'Optional column field column letters.',
            },
            filters: {
                type: 'array',
                items: { type: 'string' },
                description: 'Optional report filter column letters.',
            },
            values: {
                type: 'array',
                required: true,
                items: {
                    type: 'object',
                    additionalProperties: false,
                    properties: {
                        column: { type: 'string', description: 'Column letter of the value field.' },
                        function: { type: 'string', enum: ['sum', 'count', 'average', 'max', 'min'], required: true, description: 'Aggregation function.' },
                    },
                },
            },
            outputSheet: {
                type: 'string',
                description: 'Sheet that hosts the pivot (default "<sheet>透视").',
            },
            outPath: {
                type: 'string',
                description: 'Output .xlsx path (default: <path>.pivot.xlsx).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            const outPath = (typeof args.outPath === 'string' && args.outPath ? args.outPath : args.path.replace(/\.xlsx$/i, '.pivot.xlsx'));
            const values = args.values.map((value) => ({
                column: value.column,
                function: value.function,
            }));
            const result = await createPivotTable(args.path, {
                sheet: args.sheet,
                range: args.range,
                rows: args.rows.map(String),
                columns: Array.isArray(args.columns) ? args.columns.map(String) : undefined,
                filters: Array.isArray(args.filters) ? args.filters.map(String) : undefined,
                values,
                outputSheet: typeof args.outputSheet === 'string' ? args.outputSheet : undefined,
            }, outPath);
            return { outputPath: outPath, ...result };
        },
    })), 'tool:excel_create_pivot');
    ctx.effect(() => ctx.tools.register(defineTool({
        name: 'excel_undo',
        description: 'Undo an excel_operate edit. Restores the workbook from the snapshot taken before that edit, so everything comes back: values, formulas, styles, merged cells, row/column structure, sheet order and pivot tables. If no snapshot exists (the edit wrote a new file rather than overwriting one) it falls back to replaying the .patch.json audit log, which only restores cell content — the result says which path was used.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to the edited .xlsx file to roll back.',
            },
            patchPath: {
                type: 'string',
                required: true,
                description: 'Absolute path to the .patch.json audit log written by excel_operate.',
            },
            outPath: {
                type: 'string',
                description: 'Output .xlsx path (default: roll back in place).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            const path = args.path;
            const log = await readPatchLog(args.patchPath);
            const target = typeof args.outPath === 'string' && args.outPath ? args.outPath : path;
            // Prefer the file-level snapshot: an edit rewrites the whole package, so a
            // value-level replay cannot undo a formatting or structural change — it just
            // reports success having done nothing.
            const { restored } = await restoreSnapshot(path, target);
            if (restored) {
                return {
                    rolledBack: target,
                    method: 'snapshot',
                    patches: log.patches.length,
                    note: t('已按编辑前的完整快照恢复：单元格内容、公式、格式、合并、行列结构与透视表一并还原。'),
                };
            }
            await rollbackPatchLog(path, log, target);
            return {
                rolledBack: target,
                method: 'patch-log',
                patches: log.patches.length,
                note: t('没有找到编辑前的快照（该次编辑是写出新文件而非覆盖），只能按审计日志还原单元格内容；格式与行列结构不在日志里，无法撤销。'),
            };
        },
    })), 'tool:excel_undo');
    ctx.effect(() => ctx.tools.register(defineTool({
        name: 'excel_operate',
        description: 'Apply Excel editing operations to an .xlsx file and re-validate formulas afterwards. Operations: set (typed values/formulas), fill, fillSeries, insertRows / deleteRows / insertColumns / deleteColumns (references shift like Excel), copyRange (move:true moves, valuesOnly:true pastes cached results), transpose (paste transposed with formula shifting), copyStyle (format painter: clone font/fill/border/alignment/number format onto a range), freezeFormulas (convert formulas to their cached values), sortRange, report (one-shot report template: sort + subtotals + dynamic SUMIFS summary + filter + header style + freeze + number format), preset (role templates: ops 运营 = report + data bars; product 产品 = report + color scale; data 数分 = report + color scale + filtered copy), subtotal (group summaries), aggregateReport (dynamic pivot-style summary with live SUMIFS formulas), crosstab (two-dimension pivot grid with live SUMIFS/COUNTIFS formulas plus totals), filterToRange (advanced filter), joinSheets (exact-match two-table lookup: copy one or more columns from a lookup table back into the source by key, like VLOOKUP without formulas), fuzzyMatch (two-table fuzzy match by similarity and write the matched value back, e.g. reconcile names), uniqueValues (extract distinct values of a column), rankColumn (live RANK formula column), style, dataValidation, conditionalFormatting, autoFilter, addTable, importCsv / exportCsv (RFC 4180 with formula-injection guard), setColumnWidth / setRowHeight / autoFitColumnWidths (content-based fit, CJK aware) / freezePanes / unfreezePanes, hideRows / hideColumns, groupRows / groupColumns (outline levels, optional collapse), setZoom, showGridLines, showFormulas (saved view displays formulas instead of results), addComment (cell comments, rendered by Excel), addSparklines (per-row trend sparklines), insertImage (embed a png/jpeg/gif at a cell; cross-platform, no Excel needed), headerFooter (&-code page headers/footers), findReplace, protectSheet / unprotectSheet, mailMerge (expand {Placeholder} templates per data row), pageSetup, rowPageBreaks / clearPageBreaks (manual print breaks), printTitles (repeat header rows/columns on every printed page), definedName, setHyperlink (external URL or internal location link), addSheet / renameSheet / deleteSheet / duplicateSheet / hideSheet / setTabColor / moveSheet (reorder tabs), setWorkbookProperties (author/title/keywords + recalcOnOpen), unmergeAll, clear (cells) / clearRange (contents / formats / all), merge / unmerge, dedupeRows (remove duplicate rows by key columns, keep first/last), fillMissing (fill blanks with a value / forward from above / from the left), removeEmptyRows / removeEmptyColumns, trimText (strip whitespace), changeCase (upper / lower / proper), normalizeText (fullwidth-to-halfwidth + whitespace cleanup), splitColumn (text to columns by delimiter), highlightRows (highlight whole rows matching criteria, e.g. find and highlight a customer). The operations array is a strict union on "op": choose the matching object shape. Example: {"operations":[{"op":"set","cells":{"Sheet1!A1":"100"}},{"op":"style","range":"Sheet1!A1:C1","style":{"bold":true}}]}. Writes <path>.edited.xlsx (or outPath) and returns the post-operation validation result.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            operations: {
                type: 'array',
                items: excelOperationSchema,
                required: true,
                description: 'List of operations; each item is a strict object shape selected by its "op" field.',
            },
            outPath: {
                type: 'string',
                description: 'Output .xlsx path (default: <path>.edited.xlsx).',
            },
            validateAfter: {
                type: 'boolean',
                description: 'Re-validate the edited workbook for silent formula errors (default true).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            const path = args.path;
            const outPath = (typeof args.outPath === 'string' && args.outPath ? args.outPath : path.replace(/\.xlsx$/i, '.edited.xlsx'));
            // Validate and salvage before the executor sees them: it assumes a
            // well-formed array, so a missing nested field used to surface as a raw
            // TypeError from deep inside a handler.
            const operations = sanitizeOperations(args.operations, await readWorkbookSheetNames(path));
            const result = await operateWorkbookFile(path, operations, outPath);
            if (args.validateAfter === false) {
                return {
                    ...result,
                    validation: null,
                };
            }
            return result;
        },
    })), 'tool:excel_operate');
}
