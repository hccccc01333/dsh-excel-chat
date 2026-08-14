import { defineTool } from '@deepseek-ai/dsh-tools';
import { readFile } from 'node:fs/promises';
import { createLlmRepairAdvisor } from './advisor.js';
import { validateCharts } from './chart-validator.js';
import { exportChartsWithExcel } from './chart-visual.js';
import { readChartInfos } from './charts.js';
import { compileFormula } from './compiler.js';
import { diffWorkbookFiles } from './diff.js';
import { formulaIrSchema } from './ir-schema.js';
import { llmTextFromContext } from './llm.js';
import { operateWorkbookFile } from './operations.js';
import { repairWorkbookFile } from './repair.js';
import { detectTableFromCells } from './tables.js';
import { validate } from './validator.js';
import { visionTextFromContext } from './vision.js';
import { readWorkbookCells, validateWorkbookFile } from './workbook.js';
import { createVisionCritic } from './chart-visual.js';
export const name = 'vera-formula-validator';
export const inject = ['tools'];
export function apply(ctx) {
    console.log('[vera-formula-validator] plugin loaded');
    ctx.tools.register(defineTool({
        name: 'excel_operate',
        description: 'Apply Excel editing operations to an .xlsx file and re-validate formulas afterwards. Operations: set (typed values/formulas), fill (drag-fill with reference adjustment), insertRows / deleteRows / insertColumns / deleteColumns (references shift like Excel, including cross-sheet), addSheet / renameSheet (references update) / deleteSheet / duplicateSheet / hideSheet / setTabColor, clear, merge / unmerge, copyRange / moveRange (formulas adjust), fillSeries (numeric/date), style (bold/italic/underline/colors/numberFormat/alignment/wrap), setColumnWidth / setRowHeight / freezePanes, findReplace. Writes <path>.edited.xlsx and returns the post-operation validation result.',
        parameters: {
            path: {
                type: 'string',
                required: true,
                description: 'Absolute path to an .xlsx file.',
            },
            operations: {
                type: 'array',
                items: {
                    type: 'object',
                    additionalProperties: true,
                },
                required: true,
                description: 'List of operations, each with an "op" field (set / fill / insertRows / deleteRows / addSheet / renameSheet / deleteSheet / clear / merge / unmerge).',
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
            const outPath = args.outPath ?? args.path.replace(/\.xlsx$/i, '.edited.xlsx');
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
                return await repairWorkbookFile(args.path, advisor, cells, oracleCells);
            }
            return await repairWorkbookFile(args.path, undefined, undefined, oracleCells);
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
