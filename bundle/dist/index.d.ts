import type { Context } from '@deepseek-ai/cordis';
import Schema from '@deepseek-ai/schemastery';
export type JsonRecord = Record<string, any>;
export declare const name = "dsh-excel-chat";
export declare const inject: string[];
export interface ExcelChatConfig {
    /**
     * Language for the text a person reads: health reports, data insights, the
     * capability menu, repair summaries, operation warnings, doctor output.
     *
     * Workbook *data* is deliberately not translated — subtotal labels, generated
     * sheet names and preset names are written into the file and other code keys
     * off them (`patterns.ts` skips summary rows by matching `总计`/`小计`), so
     * translating them would silently change behaviour. The planner prompt is not
     * translated either: it is model input with tuned few-shot examples.
     */
    language?: 'zh' | 'en';
}
/**
 * dsh validates a plugin's config against this schema, so the option shows up
 * in the config surface and a typo is rejected instead of silently ignored.
 * Declared here because the loader reads `Config` from the plugin entry.
 */
export declare const Config: Schema<Schemastery.ObjectS<{
    language: Schema<"en" | "zh", "en" | "zh">;
}>, Schemastery.ObjectT<{
    language: Schema<"en" | "zh", "en" | "zh">;
}>>;
export declare function apply(host: Context, config?: ExcelChatConfig): void;
//# sourceMappingURL=index.d.ts.map