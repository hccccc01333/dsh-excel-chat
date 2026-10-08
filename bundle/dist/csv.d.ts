/**
 * RFC 4180-ish CSV parsing/writing with configurable delimiters, plus
 * formula-injection guarding for exported cells (borrowed from the
 * noatmark-dsh-plugin idea: neutralize values starting with = + - @).
 */
export declare function parseCsv(text: string, delimiter?: string): string[][];
/**
 * Neutralize spreadsheet formula injection (=, +, -, @) for literal values.
 * The prefix set follows OWASP CSV Injection guidance: Excel strips a leading
 * tab or carriage return before deciding a cell is a formula, so `\t` and `\r`
 * are attack prefixes too — not just `=`, `+`, `-`, `@`.
 */
export declare function guardFormulaInjection(value: string): string;
export declare function stringifyCsv(rows: string[][], delimiter?: string): string;
//# sourceMappingURL=csv.d.ts.map