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
/**
 * Undo `guardFormulaInjection` for a field read back from a CSV.
 *
 * The guard is one-way unless something reverses it: a text cell `=1+1` was
 * exported as `'=1+1`, and importing that CSV left the apostrophe in the cell,
 * so an export/import round-trip corrupted the value. `guarded` tells the
 * caller the field was text when it was written, which matters — letting the
 * value be re-inferred would turn it back into a live formula.
 */
export declare function unguardFormulaInjection(value: string): {
    text: string;
    guarded: boolean;
};
export declare function stringifyCsv(rows: string[][], delimiter?: string): string;
//# sourceMappingURL=csv.d.ts.map