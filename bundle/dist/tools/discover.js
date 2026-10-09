import { defineTool } from '@deepseek-ai/dsh-tools';
import { buildWorkbookMenu } from '../menu.js';
import { buildWorkbookInsight } from '../insight.js';
export function registerDiscoverTools(ctx) {
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_menu');
    ctx.effect(() => ctx.tools.register(defineTool({
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
    })), 'tool:excel_insight');
}
