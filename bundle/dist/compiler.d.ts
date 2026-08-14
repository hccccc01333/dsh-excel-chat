import type { ColumnTable, FormulaIR } from './ir.ts';
export interface CompileContext {
    baseCell: string;
    table: ColumnTable;
}
export declare function compileFormula(ir: FormulaIR, context: CompileContext): string;
//# sourceMappingURL=compiler.d.ts.map