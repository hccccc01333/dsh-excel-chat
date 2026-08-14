export type OperandIR = {
    kind: 'column';
    column: string;
} | {
    kind: 'cell';
    cell: string;
} | {
    kind: 'constant';
    value: number;
};
export interface BinaryFormulaIR {
    operation: 'binary';
    left: OperandIR;
    right: OperandIR;
    operator: '+' | '-' | '*' | '/';
}
export interface AggregateFormulaIR {
    operation: 'aggregate';
    metric: string;
    function: 'SUMIFS' | 'AVERAGEIFS' | 'COUNTIFS' | 'SUM';
    filters: Array<{
        column: string;
        value_from: string;
    }>;
}
export interface RatioFormulaIR {
    operation: 'ratio';
    numerator: OperandIR;
    denominator: OperandIR;
}
export type FormulaIR = BinaryFormulaIR | AggregateFormulaIR | RatioFormulaIR;
export interface ColumnTable {
    sheet: string;
    columns: Record<string, string>;
}
//# sourceMappingURL=ir.d.ts.map