import { defineTool } from '@deepseek-ai/dsh-tools';
import { readWorkbookCells } from '../workbook.js';
import { readFile } from 'node:fs/promises';
import { readChartInfos } from '../charts.js';
import { validateCharts } from '../chart-validator.js';
import { diffWorkbookFiles } from '../diff.js';
import { detectTableFromCells } from '../tables.js';
import { createLlmRepairAdvisor } from '../advisor.js';
import { llmTextFromContext } from '../llm.js';
import { repairWorkbookFile } from '../repair.js';
import { writeWorkbookHealthReport } from '../health-report.js';
import { autofixWorkbookFile } from '../autofix.js';
export function registerAuditTools(ctx) {
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_validate_charts');
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_diff_workbook');
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_repair_formulas');
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_health_report');
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_autofix');
}
