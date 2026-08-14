/**
 * RFC 4180-ish CSV parsing/writing with configurable delimiters, plus
 * formula-injection guarding for exported cells (borrowed from the
 * noatmark-dsh-plugin idea: neutralize values starting with = + - @).
 */
export declare function parseCsv(text: string, delimiter?: string): string[][];
/** Neutralize spreadsheet formula injection (=, +, -, @) for literal values. */
export declare function guardFormulaInjection(value: string): string;
export declare function stringifyCsv(rows: string[][], delimiter?: string): string;
//# sourceMappingURL=csv.d.ts.map