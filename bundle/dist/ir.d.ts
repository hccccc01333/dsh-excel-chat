export type OperandIR = {
    kind: 'column';
    column: string;
} | {
    kind: 'cell';
    cell: string;
} | {
    kind: 'range';
    range: string;
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
export interface FunctionFormulaIR {
    operation: 'function';
    name: 'VLOOKUP' | 'INDEX' | 'MATCH' | 'ROUND' | 'TEXT' | 'SUMIF' | 'COUNTIF' | 'AVERAGE' | 'MEDIAN' | 'MAX' | 'MIN' | 'COUNT' | 'COUNTA' | 'TODAY' | 'YEAR' | 'MONTH' | 'DAY' | 'DATE' | 'DATEDIF' | 'EOMONTH' | 'SUMIFS' | 'AVERAGEIFS' | 'COUNTIFS';
    args: OperandIR[];
}
export type FormulaIR = BinaryFormulaIR | AggregateFormulaIR | RatioFormulaIR | FunctionFormulaIR;
export interface ColumnTable {
    sheet: string;
    columns: Record<string, string>;
}
//# sourceMappingURL=ir.d.ts.map