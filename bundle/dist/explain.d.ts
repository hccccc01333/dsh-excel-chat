export interface FormulaExplanation {
    formula: string;
    summary: string;
    details: string[];
    references: string[];
}
/** Explain an Excel formula in plain language (cellm/xeli-style). */
export declare function explainFormula(formula: string): FormulaExplanation;
/** Read the formula (or value) of one cell from an .xlsx file. */
export declare function readCellContent(path: string, cellId: string): Promise<string>;
//# sourceMappingURL=explain.d.ts.map