import ExcelJS from 'exceljs';
import { type ValidationResult } from './validator.ts';
export type ExcelOperation = {
    op: 'set';
    cells: Record<string, string>;
} | {
    op: 'fill';
    source: string;
    target: string;
} | {
    op: 'insertRows';
    sheet: string;
    row: number;
    count: number;
} | {
    op: 'deleteRows';
    sheet: string;
    row: number;
    count: number;
} | {
    op: 'insertColumns';
    sheet: string;
    column: string;
    count: number;
} | {
    op: 'deleteColumns';
    sheet: string;
    column: string;
    count: number;
} | {
    op: 'addSheet';
    name: string;
} | {
    op: 'renameSheet';
    oldName: string;
    newName: string;
} | {
    op: 'deleteSheet';
    name: string;
} | {
    op: 'clear';
    cells: string[];
} | {
    op: 'merge';
    range: string;
} | {
    op: 'unmerge';
    range: string;
} | {
    op: 'copyRange';
    source: string;
    target: string;
    move?: boolean;
} | {
    op: 'fillSeries';
    start: string;
    target: string;
    step?: number;
} | {
    op: 'style';
    range: string;
    style: ExcelStyle;
} | {
    op: 'setColumnWidth';
    sheet: string;
    column: string;
    width: number;
} | {
    op: 'setRowHeight';
    sheet: string;
    row: number;
    height: number;
} | {
    op: 'freezePanes';
    sheet: string;
    row: number;
    column: string;
} | {
    op: 'findReplace';
    find: string;
    replace: string;
    sheet?: string;
    matchCase?: boolean;
} | {
    op: 'duplicateSheet';
    name: string;
    newName: string;
} | {
    op: 'hideSheet';
    name: string;
    hidden?: boolean;
} | {
    op: 'setTabColor';
    name: string;
    color: string;
} | {
    op: 'importCsv';
    file: string;
    sheet?: string;
    delimiter?: string;
    firstRowHeaders?: boolean;
} | {
    op: 'exportCsv';
    file: string;
    sheet?: string;
    range?: string;
    delimiter?: string;
    guardFormulas?: boolean;
} | {
    op: 'sortRange';
    range: string;
    keys: Array<{
        column: string;
        direction?: 'asc' | 'desc';
    }>;
    headerRows?: number;
} | {
    op: 'report';
    source: string;
    groupColumn: string;
    metrics: Array<{
        column: string;
        function: 'sum' | 'average' | 'count' | 'counta' | 'max' | 'min';
    }>;
    sort?: boolean;
    subtotal?: boolean;
    autoFilter?: boolean;
    headerStyle?: boolean;
    freezeHeader?: boolean;
    numberFormat?: string;
    outputSheet?: string;
} | {
    op: 'preset';
    role: 'ops' | 'product' | 'data';
    source: string;
    groupColumn: string;
    metrics: Array<{
        column: string;
        function: 'sum' | 'average' | 'count' | 'counta' | 'max' | 'min';
    }>;
    filter?: {
        column: string;
        operator: 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains';
        value: string | number;
    };
} | {
    op: 'dataValidation';
    range: string;
    type: 'list' | 'whole' | 'decimal' | 'date' | 'textLength' | 'custom';
    operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
    formula1?: string;
    formula2?: string;
    allowBlank?: boolean;
    showInputMessage?: boolean;
    prompt?: string;
    showErrorMessage?: boolean;
    errorStyle?: 'stop' | 'warning' | 'information';
    error?: string;
    errorTitle?: string;
} | {
    op: 'conditionalFormatting';
    range: string;
    rules: Array<{
        type: 'cellIs' | 'expression' | 'containsText' | 'notContainsText' | 'blanks' | 'noBlanks' | 'errors' | 'noErrors' | 'duplicateValues' | 'uniqueValues' | 'aboveAverage' | 'belowAverage' | 'timePeriod' | 'dataBar' | 'colorScale' | 'iconSet' | 'top10';
        operator?: string;
        formula?: string | number;
        formula2?: string | number;
        text?: string;
        timePeriod?: 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisMonth' | 'lastMonth' | 'nextMonth' | 'thisWeek' | 'lastWeek' | 'nextWeek';
        color?: string;
        minColor?: string;
        midColor?: string;
        maxColor?: string;
        iconSet?: string;
        rank?: number;
        percent?: boolean;
        bottom?: boolean;
        style?: ExcelStyle;
    }>;
} | {
    op: 'autoFilter';
    range: string;
} | {
    op: 'subtotal';
    sheet: string;
    range: string;
    groupColumn: string;
    summaryColumns: Array<{
        column: string;
        function: 'sum' | 'average' | 'count' | 'max' | 'min';
    }>;
    addGrandTotal?: boolean;
} | {
    op: 'aggregateReport';
    source: string;
    groupColumn: string;
    metrics: Array<{
        column: string;
        function: 'sum' | 'average' | 'count' | 'counta' | 'max' | 'min';
    }>;
    outputSheet?: string;
} | {
    op: 'filterToRange';
    source: string;
    criteria: Array<{
        column: string;
        operator: 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains';
        value: string | number;
    }>;
    target: string;
    matchAll?: boolean;
} | {
    op: 'protectSheet';
    sheet: string;
    password?: string;
    options?: {
        selectLockedCells?: boolean;
        selectUnlockedCells?: boolean;
        formatCells?: boolean;
        formatColumns?: boolean;
        formatRows?: boolean;
        insertColumns?: boolean;
        insertRows?: boolean;
        deleteColumns?: boolean;
        deleteRows?: boolean;
        sort?: boolean;
        autoFilter?: boolean;
    };
} | {
    op: 'unprotectSheet';
    sheet: string;
    password?: string;
} | {
    op: 'pageSetup';
    sheet: string;
    printArea?: string;
    orientation?: 'portrait' | 'landscape';
    fitToPage?: boolean;
    fitToWidth?: number;
    fitToHeight?: number;
    margins?: {
        top?: number;
        right?: number;
        bottom?: number;
        left?: number;
        header?: number;
        footer?: number;
    };
    centerHorizontally?: boolean;
    centerVertically?: boolean;
} | {
    op: 'definedName';
    name: string;
    ref: string;
} | {
    op: 'mailMerge';
    template: string;
    data: string;
    outputSheet?: string;
} | {
    op: 'addTable';
    name: string;
    range: string;
    headerRow?: boolean;
    totalsRow?: boolean;
    showRowStripes?: boolean;
    showColumnStripes?: boolean;
} | {
    op: 'dedupeRows';
    sheet: string;
    columns?: string[];
    keep?: 'first' | 'last';
} | {
    op: 'fillMissing';
    range: string;
    mode: 'value' | 'forward' | 'left';
    value?: string | number;
} | {
    op: 'removeEmptyRows';
    range: string;
} | {
    op: 'removeEmptyColumns';
    range: string;
} | {
    op: 'trimText';
    range: string;
} | {
    op: 'changeCase';
    range: string;
    case: 'upper' | 'lower' | 'proper';
} | {
    op: 'normalizeText';
    range: string;
} | {
    op: 'splitColumn';
    sheet: string;
    column: string;
    delimiter: string;
    startRow: number;
    endRow?: number;
} | {
    op: 'highlightRows';
    sheet: string;
    range: string;
    criteria: Array<{
        column: string;
        operator: 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains';
        value: string | number;
    }>;
    style?: ExcelStyle;
} | {
    op: 'fuzzyMatch';
    source: string;
    sourceKey: string;
    target: string;
    targetKey: string;
    valueColumn: string;
    outputColumn: string;
    threshold?: number;
    scoreColumn?: string;
};
export interface ExcelStyle {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontSize?: number;
    fontName?: string;
    fontColor?: string;
    fill?: string;
    numberFormat?: string;
    hAlign?: 'left' | 'center' | 'right';
    vAlign?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
    border?: BorderSpec;
}
export interface BorderEdgeSpec {
    style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
    color?: string;
}
export interface BorderSpec {
    top?: BorderEdgeSpec;
    bottom?: BorderEdgeSpec;
    left?: BorderEdgeSpec;
    right?: BorderEdgeSpec;
}
export interface OperationWarning {
    op: number;
    message: string;
}
export interface ApplyOperationsResult {
    warnings: OperationWarning[];
}
export interface OperateResult extends ApplyOperationsResult {
    outputPath: string;
    /** Path of the audit log (.patch.json) written next to the output file. */
    patchLog: string;
    validation: ValidationResult;
}
export declare function findSheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined;
/**
 * Shift selected reference points of a formula. rowDelta/colDelta apply to
 * relative rows/columns; rowThreshold/colThreshold gate the shift so row edits
 * only move references at or below the insertion/deletion point. When
 * editedSheet is set, only references pointing into that sheet are shifted.
 */
export declare function shiftFormulaReferences(formula: string, baseSheet: string, editedSheet: string | null, options?: {
    rowDelta?: number;
    colDelta?: number;
    rowThreshold?: number;
    colThreshold?: number;
    rowDeletedStart?: number;
    rowDeletedEnd?: number;
    colDeletedStart?: number;
    colDeletedEnd?: number;
}): string;
export declare function applyOperationsToWorkbook(inputPath: string, operations: ExcelOperation[], outputPath: string): Promise<ApplyOperationsResult>;
export declare function operateWorkbookFile(path: string, operations: ExcelOperation[], outputPath: string): Promise<OperateResult>;
//# sourceMappingURL=operations.d.ts.map