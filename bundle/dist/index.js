import { defineTool } from '@deepseek-ai/dsh-tools';
import { readFile } from 'node:fs/promises';
import { runAgentTask } from './agent.js';
import { createLlmRepairAdvisor } from './advisor.js';
import { autofixWorkbookFile } from './autofix.js';
import { validateCharts } from './chart-validator.js';
import { createChartWithExcel, exportChartsWithExcel, exportWorkbookToPdf, modifyChartWithExcel, } from './chart-visual.js';
import { readChartInfos } from './charts.js';
import { compileFormula } from './compiler.js';
import { diffWorkbookFiles, readPatchLog, rollbackPatchLog } from './diff.js';
import { runDoctorChecks } from './doctor.js';
import { buildWorkbookInsight } from './insight.js';
import { writeWorkbookHealthReport } from './health-report.js';
import { explainFormula, readCellContent } from './explain.js';
import { applyInPlaceEdit, revertInPlaceEdit } from './live-edit.js';
import { createLlmPlanner } from './llm-planner.js';
import { formulaIrSchema } from './ir-schema.js';
import { llmTextFromContext } from './llm.js';
import { excelOperationSchema } from './operation-schema.js';
import { operateWorkbookFile } from './operations.js';
import { buildWorkbookMenu } from './menu.js';
import { createPivotTable } from './pivot.js';
import { buildWorkbookPreview } from './preview.js';
import { profileWorkbook } from './profile.js';
import { repairWorkbookFile } from './repair.js';
import { readWorkbookDetail } from './read.js';
import { buildWorkbookSemanticProfile } from './semantic.js';
import { runExcelTask } from './task.js';
import { detectTableFromCells } from './tables.js';
import { validate } from './validator.js';
import { visionTextFromContext } from './vision.js';
import { readWorkbookCells, validateWorkbookFile } from './workbook.js';
import { createVisionCritic } from './chart-visual.js';
export const name = 'vera-formula-validator';
export const inject = ['tools', 'systemPrompt'];
export function apply(ctx) {
    console.log('[vera-formula-validator] plugin loaded');
    const commands = ctx.get('commands');
    if (commands) {
        commands.register({
            name: 'excel-set',
            description: '就地修改本地 Excel 文件的一个单元格（自动备份 .bak + 审计日志 + 公式体检）',
            input: { hint: '{"path":"D:\\\\x.xlsx","cell":"Sheet1!A1","value":"..."}' },
            handler: async (invocation) => {
                try {
                    const payload = JSON.parse(invocation.rawInput.trim());
                    if (!payload.path || !payload.cell || payload.value === undefined) {
                        return { kind: 'error', text: '需要 {"path","cell","value"}' };
                    }
                    const result = await applyInPlaceEdit(payload.path, payload.cell, payload.value);
                    return {
                        kind: 'success',
                        text: `已就地保存 ${result.cell} = ${result.value}（备份 ${result.backupPath}，公式异常 ${result.anomalies}）`,
                    };
                }
                catch (error) {
                    return { kind: 'error', text: error instanceof Error ? error.message : String(error) };
                }
            },
        });
        commands.register({
            name: 'excel-undo',
            description: '回滚本地 Excel 文件最近一次就地修改',
            input: { hint: '{"path":"D:\\\\x.xlsx"}' },
            handler: async (invocation) => {
                try {
                    const payload = JSON.parse(invocation.rawInput.trim());
                    if (!payload.path)
                        return { kind: 'error', text: '需要 {"path"}' };
                    const outcome = await revertInPlaceEdit(payload.path);
                    return { kind: outcome.restored ? 'success' : 'error', text: outcome.message };
                }
                catch (error) {
                    return { kind: 'error', text: error instanceof Error ? error.message : String(error) };
                }
            },
        });
        commands.register({
            name: 'excel-doctor',
            description: '安装自检：检查宿主包隔离、Node 版本与 Excel 引擎冒烟',
            input: { hint: '可选：{"profile":"D:\\\\profile目录"}' },
            handler: async (invocation) => {
                try {
                    const payload = JSON.parse(invocation.rawInput.trim() || '{}');
                    const checks = await runDoctorChecks({ profileDirs: payload.profile ? [payload.profile] : [] });
                    const lines = checks.map((check) => `${check.ok ? '✅' : '❌'} ${check.name}: ${check.detail}`);
                    const failed = checks.filter((check) => !check.ok).length;
                    return { kind: failed === 0 ? 'success' : 'error', text: lines.join('\n') };
                }
                catch (error) {
                    return { kind: 'error', text: error instanceof Error ? error.message : String(error) };
                }
            },
        });
    }
    ctx.systemPrompt.section({
        name: 'dsh-excel-chat:interaction',
        order: 150,
        text: [
            'Excel 对话交互原则：',
            '- 用户说业务目标而不是操作时（例如“做周报”“帮我整理一下”），不要追问技术细节：调用 excel_menu 给出 2-3 个可选方案让用户挑。',
            '- 用户给出文件但没说要做什么时，先调用 excel_profile 或 excel_menu，主动介绍文件里有什么、列出能做的事，让用户选择。',
            '- 能合理猜出意图时，直接做最可能的版本并展示结果，说明不满意可以用 excel_undo 回滚；删除行列、覆盖数据、删除/保护工作表等破坏性操作必须先确认。',
            '- 运营 / 产品 / 数分岗位用户可以直接套 preset 岗位模板。',
        ].join('\n'),
    });
    ctx.tools.register(defineTool({
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
    }));
    ctx.tools.register(defineTool({
        name: 'excel_undo',
        description: 'Roll back an excel_operate edit using its .patch.json audit log: every changed cell is restored to the pre-edit state. Content-level undo (inserted/deleted rows and columns are not removed, but their cells are restored).',
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
            const log = await readPatchLog(args.patchPath);
            const target = typeof args.outPath === 'string' && args.outPath ? args.outPath : args.path;
            await rollbackPatchLog(args.path, log, target);
            return { rolledBack: target, patches: log.patches.length };
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_operate',
        description: 'Apply Excel editing operations to an .xlsx file and re-validate formulas afterwards. Operations: set (typed values/formulas), fill, fillSeries, insertRows / deleteRows / insertColumns / deleteColumns (references shift like Excel), copyRange (move:true moves, valuesOnly:true pastes cached results), transpose (paste transposed with formula shifting), copyStyle (format painter: clone font/fill/border/alignment/number format onto a range), freezeFormulas (convert formulas to their cached values), sortRange, report (one-shot report template: sort + subtotals + dynamic SUMIFS summary + filter + header style + freeze + number format), preset (role templates: ops 运营 = report + data bars; product 产品 = report + color scale; data 数分 = report + color scale + filtered copy), subtotal (group summaries), aggregateReport (dynamic pivot-style summary with live SUMIFS formulas), crosstab (two-dimension pivot grid with live SUMIFS/COUNTIFS formulas plus totals), filterToRange (advanced filter), joinSheets (exact-match two-table lookup: copy one or more columns from a lookup table back into the source by key, like VLOOKUP without formulas), fuzzyMatch (two-table fuzzy match by similarity and write the matched value back, e.g. reconcile names), uniqueValues (extract distinct values of a column), rankColumn (live RANK formula column), style, dataValidation, conditionalFormatting, autoFilter, addTable, importCsv / exportCsv (RFC 4180 with formula-injection guard), setColumnWidth / setRowHeight / autoFitColumnWidths (content-based fit, CJK aware) / freezePanes / unfreezePanes, hideRows / hideColumns, groupRows / groupColumns (outline levels, optional collapse), setZoom, showGridLines, addComment (cell comments, rendered by Excel), addSparklines (per-row trend sparklines), headerFooter (&-code page headers/footers), findReplace, protectSheet / unprotectSheet, mailMerge (expand {Placeholder} templates per data row), pageSetup, rowPageBreaks / clearPageBreaks (manual print breaks), printTitles (repeat header rows/columns on every printed page), definedName, setHyperlink (external URL or internal location link), addSheet / renameSheet / deleteSheet / duplicateSheet / hideSheet / setTabColor / moveSheet (reorder tabs), setWorkbookProperties (author/title/keywords + recalcOnOpen), unmergeAll, clear (cells) / clearRange (contents / formats / all), merge / unmerge, dedupeRows (remove duplicate rows by key columns, keep first/last), fillMissing (fill blanks with a value / forward from above / from the left), removeEmptyRows / removeEmptyColumns, trimText (strip whitespace), changeCase (upper / lower / proper), normalizeText (fullwidth-to-halfwidth + whitespace cleanup), splitColumn (text to columns by delimiter), highlightRows (highlight whole rows matching criteria, e.g. find and highlight a customer). The operations array is a strict union on "op": choose the matching object shape. Example: {"operations":[{"op":"set","cells":{"Sheet1!A1":"100"}},{"op":"style","range":"Sheet1!A1:C1","style":{"bold":true}}]}. Writes <path>.edited.xlsx (or outPath) and returns the post-operation validation result.',
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
            const outPath = (typeof args.outPath === 'string' && args.outPath ? args.outPath : args.path.replace(/\.xlsx$/i, '.edited.xlsx'));
            const result = await operateWorkbookFile(args.path, args.operations, outPath);
            if (args.validateAfter === false) {
                return {
                    ...result,
                    validation: null,
                };
            }
            return result;
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_read',
        description: 'Precisely read cells from an .xlsx file: values, formulas, value types, number formats, font/fill/alignment, merged ranges, and data validation. Use before editing to inspect the exact cell state. For large sheets, run excel_profile first, then read page by page with range + maxRows (e.g. range "A1:E101" maxRows 100, then "A102:E201" maxRows 100).',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            sheet: {
                type: 'string',
                description: 'Restrict to one sheet (default all sheets).',
            },
            range: {
                type: 'string',
                description: 'A1 range on the selected sheet, e.g. "A1:D20".',
            },
            cells: {
                type: 'array',
                items: { type: 'string' },
                description: 'Exact cell ids to read, e.g. ["A1","D4"].',
            },
            maxRows: {
                type: 'number',
                description: 'Cap the number of rows read (the result marks truncated). Combine with a range start row for paging through large sheets.',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            const cells = Array.isArray(args.cells) ? args.cells.map((cell) => String(cell)) : undefined;
            const sheets = await readWorkbookDetail(args.path, {
                sheet: typeof args.sheet === 'string' ? args.sheet : undefined,
                range: typeof args.range === 'string' ? args.range : undefined,
                cells,
                maxRows: typeof args.maxRows === 'number' && args.maxRows > 0 ? args.maxRows : undefined,
            });
            // dsh requires lossless JSON: strip optional undefined fields explicitly.
            return JSON.parse(JSON.stringify({ path: args.path, sheets }));
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_profile',
        description: 'Compact structural digest of an .xlsx file: per-sheet dimensions, detected header row, formula-cell count, and per-column dtype/missing/unique counts, numeric min/max/mean, date range, top values, and samples, plus the range to read first. Use this before excel_read on large or unfamiliar files so the conversation does not dump whole sheets.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            sheet: {
                type: 'string',
                description: 'Restrict to one sheet (default all sheets).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            return await profileWorkbook(args.path, typeof args.sheet === 'string' ? args.sheet : undefined);
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_semantic_profile',
        description: 'Workbook semantic layer: classify every column as time / measure / dimension / id, derive each sheet grain and formula-based metrics, and find cross-sheet join keys. Use before analysis tasks ("按区域汇总金额", "哪个地区利润下降最严重") so the agent knows which columns are dimensions/measures instead of guessing.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            sheet: {
                type: 'string',
                description: 'Restrict to one sheet (default all sheets).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            return await buildWorkbookSemanticProfile(args.path, typeof args.sheet === 'string' ? args.sheet : undefined);
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_preview',
        description: 'Human-facing table preview: render the requested sheet/range as a Markdown table (embed it in your reply so the user sees the data inline) and write an HTML preview file next to the workbook. Call when the user asks to 看看/展示/预览 the table or wants to see what a sheet looks like before deciding next steps.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            sheet: {
                type: 'string',
                description: 'Sheet name (default first sheet).',
            },
            range: {
                type: 'string',
                description: 'A1 range to preview, e.g. "A1:F20".',
            },
            maxRows: {
                type: 'number',
                description: 'Cap preview rows (default 20).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            return await buildWorkbookPreview(args.path, {
                sheet: typeof args.sheet === 'string' ? args.sheet : undefined,
                range: typeof args.range === 'string' ? args.range : undefined,
                maxRows: typeof args.maxRows === 'number' && args.maxRows > 0 ? args.maxRows : undefined,
            });
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_menu',
        description: 'Turn an .xlsx file into a ready-made menu: profile it, summarize in plain language what the file contains, and list concrete next steps (clean / fill missing / health-check / report / pivot / chart / mail merge / role preset), each with an example prompt the user can just confirm. Call this when the user gives a file but has not said what to do, or when their request is a business goal ("make me a weekly report") instead of a concrete operation. Present the menu and let the user pick a number or send the example; do not ask open technical questions.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            sheet: {
                type: 'string',
                description: 'Restrict to one sheet (default all sheets).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            return await buildWorkbookMenu(args.path, typeof args.sheet === 'string' ? args.sheet : undefined);
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_insight',
        description: 'Data insight report for an .xlsx file: per-sheet plain-language summary plus heuristic anomaly findings (missing values, suspicious duplicates, outlier/negative values, text whitespace, formula presence) and concrete next-step suggestions. Call when the user asks "summarize this file", "帮我看看这表有什么问题", or wants to know what the data says before doing anything.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            sheet: {
                type: 'string',
                description: 'Restrict to one sheet (default all sheets).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            return await buildWorkbookInsight(args.path, typeof args.sheet === 'string' ? args.sheet : undefined);
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_validate_charts_visual',
        description: 'Export all charts from an .xlsx file to PNG using local Excel, then ask the configured vision-capable LLM to check visual quality (title, legend, labels, axes, crowding, trend readability).',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            model: {
                type: 'string',
                required: true,
                description: 'Vision-capable LLM model id.',
            },
            provider: {
                type: 'string',
                description: 'LLM provider route (default "deepseek").',
            },
            outDir: {
                type: 'string',
                description: 'Output directory for PNG files (default: <path>.charts).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args, exec) {
            const outDir = args.outDir ?? `${args.path.replace(/\.xlsx$/i, '')}.charts`;
            const images = await exportChartsWithExcel(args.path, outDir, exec.signal);
            const critic = createVisionCritic(visionTextFromContext(ctx, args.provider ?? 'deepseek', args.model));
            const reports = [];
            for (const image of images) {
                reports.push(await critic(image, exec.signal));
            }
            return { images, reports };
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_create_chart',
        description: 'Create a chart in an .xlsx copy using local Excel (Windows only): choose data range, chart type (column/line/pie/bar/area), title, and chart name.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            range: {
                type: 'string',
                required: true,
                description: 'Data range including headers, e.g. "Sheet1!A1:B4".',
            },
            sheet: {
                type: 'string',
                description: 'Sheet name containing the range (default first sheet).',
            },
            type: {
                type: 'string',
                enum: ['column', 'line', 'pie', 'bar', 'area'],
                description: 'Chart type (default column).',
            },
            title: {
                type: 'string',
                description: 'Chart title.',
            },
            name: {
                type: 'string',
                description: 'Chart name (default "Chart 1").',
            },
            outPath: {
                type: 'string',
                description: 'Output .xlsx path (default: <path>.chart.xlsx).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args, exec) {
            if (process.platform !== 'win32') {
                throw new Error('chart creation requires Windows with Microsoft Excel installed');
            }
            const outPath = (typeof args.outPath === 'string' && args.outPath ? args.outPath : args.path.replace(/\.xlsx$/i, '.chart.xlsx'));
            await createChartWithExcel(args.path, {
                sheet: typeof args.sheet === 'string' ? args.sheet : undefined,
                range: args.range,
                type: args.type,
                title: typeof args.title === 'string' ? args.title : undefined,
                name: typeof args.name === 'string' ? args.name : undefined,
            }, outPath, exec.signal);
            return { outputPath: outPath };
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_modify_chart',
        description: 'Modify chart parameters in an .xlsx copy using local Excel (Windows only): chart type, title, legend, and axis titles.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            chartName: {
                type: 'string',
                required: true,
                description: 'Name of the chart to modify (e.g. "Chart 1").',
            },
            type: {
                type: 'string',
                enum: ['column', 'line', 'pie', 'bar', 'area'],
                description: 'New chart type.',
            },
            title: {
                type: 'string',
                description: 'New chart title.',
            },
            hasLegend: {
                type: 'boolean',
                description: 'Show or hide the legend.',
            },
            axisTitleX: {
                type: 'string',
                description: 'Category axis title.',
            },
            axisTitleY: {
                type: 'string',
                description: 'Value axis title.',
            },
            outPath: {
                type: 'string',
                description: 'Output .xlsx path (default: <path>.chart.xlsx).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args, exec) {
            if (process.platform !== 'win32') {
                throw new Error('chart modification requires Windows with Microsoft Excel installed');
            }
            const outPath = (typeof args.outPath === 'string' && args.outPath ? args.outPath : args.path.replace(/\.xlsx$/i, '.chart.xlsx'));
            await modifyChartWithExcel(args.path, args.chartName, {
                type: args.type,
                title: typeof args.title === 'string' ? args.title : undefined,
                hasLegend: typeof args.hasLegend === 'boolean' ? args.hasLegend : undefined,
                axisTitleX: typeof args.axisTitleX === 'string' ? args.axisTitleX : undefined,
                axisTitleY: typeof args.axisTitleY === 'string' ? args.axisTitleY : undefined,
            }, outPath, exec.signal);
            return { outputPath: outPath };
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_export_charts',
        description: 'Export all charts from an .xlsx file to PNG images using local Microsoft Excel (Windows only).',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            outDir: {
                type: 'string',
                description: 'Output directory for PNG files (default: <path>.charts).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args, exec) {
            const outDir = args.outDir ?? `${args.path.replace(/\.xlsx$/i, '')}.charts`;
            return { images: await exportChartsWithExcel(args.path, outDir, exec.signal) };
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_export_pdf',
        description: 'Export an .xlsx file (whole workbook or one sheet) to a PDF using local Microsoft Excel COM. Windows only; the source file is opened read-only and left untouched.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            outPath: {
                type: 'string',
                description: 'Output PDF path (default: <path>.pdf).',
            },
            sheet: {
                type: 'string',
                description: 'Export only this sheet (default all sheets).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args, exec) {
            const outPath = typeof args.outPath === 'string' && args.outPath
                ? args.outPath
                : args.path.replace(/\.xlsx$/i, '.pdf');
            await exportWorkbookToPdf(args.path, outPath, typeof args.sheet === 'string' && args.sheet ? args.sheet : undefined, exec.signal);
            return { pdfPath: outPath };
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_validate_charts',
        description: 'Validate chart structure inside an .xlsx file: chart type, series references, missing cells, two-dimensional ranges, and unsorted date categories.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            const cells = await readWorkbookCells(await readFile(args.path));
            const charts = await readChartInfos(args.path);
            return { charts, reports: validateCharts(charts, cells) };
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_diff_workbook',
        description: 'Compare two .xlsx files cell by cell and return added/removed/changed cells. Useful as a workbook git diff.',
        parameters: {
            beforePath: {
                type: 'string',
                required: true,
                description: 'Absolute path to the original .xlsx file.',
            },
            afterPath: {
                type: 'string',
                required: true,
                description: 'Absolute path to the modified .xlsx file.',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            return { entries: await diffWorkbookFiles(args.beforePath, args.afterPath) };
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_repair_formulas',
        description: 'Validate an .xlsx file, generate deterministic repairs for reference-pattern anomalies, optionally ask an LLM to repair remaining anomalies via Formula IR, write a .repaired.xlsx copy, and re-validate.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            useLlm: {
                type: 'boolean',
                description: 'Ask the configured LLM to repair anomalies the deterministic generator cannot fix.',
            },
            autoTable: {
                type: 'boolean',
                description: 'Detect the header row from cell content when no table schema is provided (uses the first row with two or more text cells).',
            },
            provider: {
                type: 'string',
                description: 'LLM provider route (default "deepseek").',
            },
            model: {
                type: 'string',
                description: 'LLM model id. Required when useLlm is true.',
            },
            table: {
                type: 'object',
                additionalProperties: true,
                description: 'Table schema { sheet, columns } for LLM repair compilation. Required when useLlm is true unless autoTable is enabled.',
            },
            oraclePath: {
                type: 'string',
                description: 'Absolute path to the ground-truth .xlsx file. When provided, the repaired workbook is scored against it and the result includes oracleScore.',
            },
            outPath: {
                type: 'string',
                description: 'Output .xlsx path (default: <path>.repaired.xlsx).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args, exec) {
            let oracleCells;
            if (args.oraclePath) {
                oracleCells = await readWorkbookCells(await readFile(args.oraclePath));
            }
            const outputPath = typeof args.outPath === 'string' && args.outPath ? args.outPath : undefined;
            if (args.useLlm) {
                if (!args.model) {
                    throw new Error('model is required when useLlm is true');
                }
                let table = args.table;
                let cells;
                if (!table) {
                    if (!args.autoTable) {
                        throw new Error('table (or autoTable: true) is required when useLlm is true');
                    }
                    cells = await readWorkbookCells(await readFile(args.path));
                    const detected = detectTableFromCells(cells);
                    if (!detected) {
                        throw new Error('autoTable could not detect a header row; provide table explicitly');
                    }
                    table = detected;
                }
                const advisor = createLlmRepairAdvisor(llmTextFromContext(ctx, args.provider ?? 'deepseek', args.model), table, exec.signal);
                return await repairWorkbookFile(args.path, advisor, cells, oracleCells, outputPath);
            }
            return await repairWorkbookFile(args.path, undefined, undefined, oracleCells, outputPath);
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_health_report',
        description: 'Write a formula health report INTO the workbook itself: a hidden `_dsh_体检报告` sheet with a health score (100 minus 10 per anomaly), formula/anomaly counts, and one row per anomaly (cell/kind/reason). The file carries its own audit trail. Call after excel_autofix / excel_task or when the user wants the report to travel with the file.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            outPath: {
                type: 'string',
                description: 'Output .xlsx path (default: write in place).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            return await writeWorkbookHealthReport(args.path, typeof args.outPath === 'string' && args.outPath ? args.outPath : undefined);
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_autofix',
        description: 'One-call self-healing loop for an .xlsx file: validate formulas, apply deterministic repairs for reference-pattern anomalies, optionally ask an LLM to repair the rest via Formula IR, re-validate the repaired copy, and report a plain-language before/after summary. Use after edits or when the user asks to "check and fix" a workbook.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            useLlm: {
                type: 'boolean',
                description: 'Ask the configured LLM to repair anomalies the deterministic generator cannot fix.',
            },
            autoTable: {
                type: 'boolean',
                description: 'Detect the header row from cell content when no table schema is provided.',
            },
            provider: {
                type: 'string',
                description: 'LLM provider route (default "deepseek").',
            },
            model: {
                type: 'string',
                description: 'LLM model id. Required when useLlm is true.',
            },
            table: {
                type: 'object',
                additionalProperties: true,
                description: 'Table schema { sheet, columns } for LLM repair compilation. Required when useLlm is true unless autoTable is enabled.',
            },
            outPath: {
                type: 'string',
                description: 'Output .xlsx path (default: <path>.repaired.xlsx).',
            },
            healthReport: {
                type: 'boolean',
                description: 'Embed a hidden `_dsh_体检报告` sheet into the repaired file (default true).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args, exec) {
            const outputPath = typeof args.outPath === 'string' && args.outPath ? args.outPath : undefined;
            if (!args.useLlm) {
                return await autofixWorkbookFile(args.path, {
                    outPath: outputPath,
                    healthReport: args.healthReport !== false,
                });
            }
            if (!args.model) {
                throw new Error('model is required when useLlm is true');
            }
            let table = args.table;
            let cells;
            if (!table) {
                if (!args.autoTable) {
                    throw new Error('table (or autoTable: true) is required when useLlm is true');
                }
                cells = await readWorkbookCells(await readFile(args.path));
                const detected = detectTableFromCells(cells);
                if (!detected) {
                    throw new Error('autoTable could not detect a header row; provide table explicitly');
                }
                table = detected;
            }
            const advisor = createLlmRepairAdvisor(llmTextFromContext(ctx, args.provider ?? 'deepseek', args.model), table, exec.signal);
            return await autofixWorkbookFile(args.path, {
                advisor,
                outPath: outputPath,
                healthReport: args.healthReport !== false,
            });
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_task',
        description: 'Execute an Excel workflow in one call. Two modes: (1) steps mode - provide an ordered array of excel_operate-style step operations; after every step the formulas are validated and deterministic repairs are applied, then the next step runs on the verified result. (2) goal mode - provide a natural-language goal and the configured LLM plans the steps, executes them with verification, an LLM verifier checks the goal, and it replans up to maxRounds times until achieved. Use goal mode for vague requests like "make this a monthly report" and steps mode for concrete pipelines like "clean, fill missing, then summarize by region".',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            steps: {
                type: 'array',
                required: true,
                items: {
                    type: 'object',
                    additionalProperties: false,
                    properties: {
                        name: { type: 'string', description: 'Step label, e.g. "clean".' },
                        operations: {
                            type: 'array',
                            required: true,
                            items: { type: 'object', additionalProperties: true },
                            description: 'excel_operate-style operations for this step (see excel_operate).',
                        },
                        verify: { type: 'boolean', description: 'Validate and auto-repair formulas after this step (default true).' },
                    },
                },
                description: 'Ordered steps; each runs on the previous verified result.',
            },
            goal: {
                type: 'string',
                description: 'Natural-language goal (goal mode). Exactly one of steps or goal must be provided.',
            },
            maxRounds: {
                type: 'number',
                description: 'Maximum plan-execute-verify-replan rounds in goal mode (default 2).',
            },
            provider: {
                type: 'string',
                description: 'LLM provider route for goal mode (default "deepseek").',
            },
            model: {
                type: 'string',
                description: 'LLM model id for goal mode. Required when goal is provided.',
            },
            outPath: {
                type: 'string',
                description: 'Output .xlsx path (default: <path>.task.xlsx).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            const hasSteps = Array.isArray(args.steps);
            const hasGoal = typeof args.goal === 'string' && args.goal.length > 0;
            if (hasSteps === hasGoal) {
                throw new Error('exactly one of steps or goal must be provided');
            }
            const outPath = typeof args.outPath === 'string' && args.outPath ? args.outPath : undefined;
            if (hasSteps) {
                return await runExcelTask(args.path, args.steps, outPath);
            }
            if (!args.model) {
                throw new Error('model is required when goal is provided');
            }
            const planner = createLlmPlanner(llmTextFromContext(ctx, args.provider ?? 'deepseek', args.model));
            return await runAgentTask(args.path, {
                goal: args.goal,
                planner,
                maxRounds: typeof args.maxRounds === 'number' && args.maxRounds > 0 ? args.maxRounds : 2,
                outPath,
            });
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_explain_formula',
        description: 'Explain an Excel formula in plain language: parsed functions (SUMIFS / VLOOKUP / IF / date / text / statistics), referenced ranges, cross-sheet references, and arithmetic/comparison. Pass the formula directly, or a path + cell to read it from a workbook. Use when the user asks "这个公式是什么意思".',
        parameters: {
            formula: {
                type: 'string',
                description: 'Formula text, e.g. "=VLOOKUP(A2,Sheet2!$A$1:$B$100,2,FALSE)". Exactly one of formula or path+cell must be provided.',
            },
            path: {
                type: 'string',
                description: 'Absolute path to an .xlsx file (with cell).',
            },
            cell: {
                type: 'string',
                description: 'Cell id to read, e.g. "Sheet1!D4" (with path).',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            const hasFormula = typeof args.formula === 'string' && args.formula.length > 0;
            const hasCell = typeof args.path === 'string' && typeof args.cell === 'string';
            if (hasFormula === hasCell) {
                throw new Error('exactly one of formula or path+cell must be provided');
            }
            const formula = hasFormula ? args.formula : await readCellContent(args.path, args.cell);
            if (!formula.trim().startsWith('=')) {
                throw new Error(`cell is not a formula: ${args.cell}`);
            }
            return explainFormula(formula);
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_compile_formula',
        description: 'Compile a semantic Formula IR into a deterministic Excel formula. IR operations: binary (left/right operands with operator), ratio (numerator/denominator), aggregate (metric, function, filters with value_from cell/column/constant). Table schema: { sheet, columns: { logicalName: columnLetter } }.',
        parameters: {
            ir: {
                ...formulaIrSchema,
                required: true,
                description: 'Formula IR object.',
            },
            baseCell: {
                type: 'string',
                required: true,
                description: 'Cell id where the formula will be placed (e.g. "B2" or "Sheet1!B2").',
            },
            table: {
                type: 'object',
                additionalProperties: true,
                required: true,
                description: 'Table schema mapping logical column names to column letters.',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            return {
                formula: compileFormula(args.ir, {
                    baseCell: args.baseCell,
                    table: args.table,
                }),
            };
        },
    }));
    ctx.tools.register(defineTool({
        name: 'excel_validate_formulas',
        description: 'Scan workbook cells for silent formula errors: inconsistent reference patterns inside a column, hardcoded values in formula columns, empty fill gaps, and circular references. Pass cells as an object mapping cell id (e.g. "Sheet1!D4" or "D4") to cell content (formulas start with "=").',
        parameters: {
            path: {
                type: 'string',
                description: 'Absolute path to an .xlsx file. Exactly one of path or cells must be provided.',
            },
            cells: {
                type: 'object',
                additionalProperties: true,
                description: 'Map of cell id to cell content. Exactly one of path or cells must be provided.',
            },
        },
        output: {
            schema: { type: 'object', additionalProperties: true },
            render: (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }],
        },
        async execute(args) {
            const hasPath = typeof args.path === 'string' && args.path.length > 0;
            const hasCells = args.cells !== undefined;
            if (hasPath === hasCells) {
                throw new Error('exactly one of path or cells must be provided');
            }
            if (hasPath) {
                return await validateWorkbookFile(args.path);
            }
            const normalized = {};
            for (const [id, content] of Object.entries(args.cells)) {
                normalized[id] = typeof content === 'string' ? content : String(content);
            }
            return validate(normalized);
        },
    }));
}
