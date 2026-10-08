/**
 * Language for the text a person reads: health reports, insights, the
 * capability menu, repair summaries, operation warnings, doctor output.
 *
 * The plugin's output is authored in Chinese and translated here. A deployment
 * switches it with the plugin's `language` config option (see the README) rather
 * than hard-coding the choice at every call site.
 *
 * Two decisions worth knowing before adding an entry:
 *
 * 1. **The catalog is keyed by the Chinese source text**, not by invented keys.
 *    Call sites stay readable, and a string nobody has translated yet degrades
 *    to its Chinese original instead of to a missing-key marker. Interpolated
 *    values become `{name}` placeholders so one entry serves every call.
 *
 * 2. **Workbook *data* is out of scope, only *messages* are translated.**
 *    Subtotal labels (`总计`), generated sheet names (`-汇总`) and preset names
 *    are written into the file, not shown as text, and other code keys off them
 *    — `patterns.ts` skips summary rows by matching `/^(总计|小计)$/`, so
 *    translating that label would silently stop the check from skipping them.
 *    The same goes for the planner prompt in `llm-planner.ts`: it is model
 *    input with tuned few-shot examples, and the model reads both languages.
 */
export type Language = 'zh' | 'en';
/** Set the language for user-facing messages. Called once from the plugin config. */
export declare function setLanguage(next: Language): void;
export declare function getLanguage(): Language;
/**
 * Translate a user-facing message, substituting `{name}` placeholders. Falls
 * back to the Chinese source when no translation is registered, so a missing
 * entry shows the original wording rather than a key.
 */
export declare function t(zh: string, params?: Record<string, string | number>): string;
/** True when an English translation is registered for this source text. */
export declare function hasTranslation(zh: string): boolean;
/** Source strings that have no English translation yet, for the coverage test. */
export declare function untranslated(keys: readonly string[]): string[];
//# sourceMappingURL=i18n.d.ts.map