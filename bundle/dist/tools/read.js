import { defineTool } from '@deepseek-ai/dsh-tools';
import { readWorkbookDetail } from '../read.js';
import { profileWorkbook } from '../profile.js';
import { buildWorkbookSemanticProfile } from '../semantic.js';
import { buildWorkbookPreview } from '../preview.js';
export function registerReadTools(ctx) {
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_read');
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_profile');
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_semantic_profile');
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_preview');
}
