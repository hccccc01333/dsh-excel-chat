/**
 * Appearance: styles, data validation, conditional formatting and real tables.
 *
 * ExcelJS and the plugin describe styles differently (hAlign vs horizontal, fill as
 * a colour vs a pattern), so the conversion lives here in one place rather than at
 * each call site. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs';
import type { ExcelOperation, ExcelStyle } from '../operation-types.ts';
export declare function applyDataValidation(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'dataValidation';
}>): void;
export declare function looksLikeRange(value: string): boolean;
export declare function applyConditionalFormatting(workbook: ExcelJS.Workbook, range: string, rules: Extract<ExcelOperation, {
    op: 'conditionalFormatting';
}>['rules']): void;
/**
 * Map the plugin's style vocabulary onto an ExcelJS style, merged over what is there.
 *
 * One mapper for both callers: `applyStyle` and the `style` of a conditional-formatting
 * rule. They used to have a version each, and the rule's version handled five of the
 * seventeen fields — so a rule asking for `numberFormat`, `hAlign`, `wrapText` or a
 * border silently produced an empty `<dxf/>` and the user's formatting vanished.
 *
 * `fillTarget` is not a wart, it is OOXML: a conditional-format `dxf` paints the
 * background (`bgColor`) while a normal cell fill paints the foreground (`fgColor`).
 * Writing `fgColor` in a rule renders nothing, and the plugin's own test caught exactly
 * that when this mapper briefly used one target for both.
 *
 * Unspecified properties are left as `current` has them rather than cleared, so a
 * partial style adds to a cell instead of replacing its appearance.
 */
export declare function excelStylePatch(style: ExcelStyle, current?: Partial<ExcelJS.Style>, fillTarget?: 'fgColor' | 'bgColor'): Partial<ExcelJS.Style>;
export declare function excelStyleToWorkbookStyle(style: ExcelStyle): ExcelJS.Style;
export declare function addTable(workbook: ExcelJS.Workbook, options: Extract<ExcelOperation, {
    op: 'addTable';
}>): void;
export declare function applyStyle(workbook: ExcelJS.Workbook, range: string, style: ExcelStyle): void;
export declare function normalizeColor(color: string): string;
//# sourceMappingURL=formatting.d.ts.map