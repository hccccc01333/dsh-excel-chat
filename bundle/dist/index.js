import Schema from '@deepseek-ai/schemastery';
import { setLanguage } from './i18n.js';
import { registerCommands } from './tools/commands.js';
import { registerWriteTools } from './tools/write.js';
import { registerReadTools } from './tools/read.js';
import { registerDiscoverTools } from './tools/discover.js';
import { registerChartTools } from './tools/charts.js';
import { registerAuditTools } from './tools/audit.js';
import { registerFormulaTools } from './tools/formulas.js';
import { announce, describeError, guardedContext, registrationSummary, writeStatusReport } from './registration.js';
export const name = 'dsh-excel-chat';
export const inject = ['tools', 'systemPrompt'];
/**
 * dsh validates a plugin's config against this schema, so the option shows up
 * in the config surface and a typo is rejected instead of silently ignored.
 * Declared here because the loader reads `Config` from the plugin entry.
 */
export const Config = Schema.object({
    language: Schema.union(['zh', 'en']).default('zh'),
});
export function apply(host, config) {
    // Read the config defensively: the loader passes whatever the entry declared,
    // and an unrecognised value must fall back to Chinese rather than throw.
    setLanguage(config?.language === 'en' ? 'en' : 'zh');
    const { ctx, report } = guardedContext(host);
    announce(host, 'info', '[dsh-excel-chat] plugin loaded');
    try {
        registerAll(ctx);
    }
    catch (error) {
        // Individual registrations are isolated inside `guardedContext`; reaching
        // here means something outside them broke, which would otherwise look
        // exactly like the plugin never loading.
        announce(host, 'error', `[dsh-excel-chat] apply aborted: ${describeError(error)}`);
        throw error;
    }
    announce(host, report.failures.length === 0 ? 'info' : 'error', registrationSummary(report));
    writeStatusReport(report);
}
/**
 * Every registration the plugin makes.
 *
 * Kept out of `apply` so that `apply` can report a failure of the registration
 * path itself, separately from a failure of one individual tool.
 */
function registerAll(ctx) {
    registerCommands(ctx);
    ctx.effect(() => ctx.systemPrompt.section({
        name: 'dsh-excel-chat:interaction',
        order: 150,
        text: [
            'Excel 对话交互原则：',
            '- 用户说业务目标而不是操作时（例如“做周报”“帮我整理一下”），不要追问技术细节：调用 excel_menu 给出 2-3 个可选方案让用户挑。',
            '- 用户给出文件但没说要做什么时，先调用 excel_profile 或 excel_menu，主动介绍文件里有什么、列出能做的事，让用户选择。',
            '- 能合理猜出意图时，直接做最可能的版本并展示结果，说明不满意可以用 excel_undo 回滚；删除行列、覆盖数据、删除/保护工作表等破坏性操作必须先确认。',
            '- 运营 / 产品 / 数分岗位用户可以直接套 preset 岗位模板。',
        ].join('\n'),
    }), 'system-prompt:dsh-excel-chat');
    registerWriteTools(ctx);
    registerReadTools(ctx);
    registerDiscoverTools(ctx);
    registerChartTools(ctx);
    registerAuditTools(ctx);
    registerFormulaTools(ctx);
}
