/**
 * Strict dsh tool-DSL schema for excel_operate operations. Each operation is a
 * discriminated union on `op`, so the model knows the exact fields to emit
 * instead of guessing from prose.
 */
const text = (description: string, required = false) => ({ type: 'string', description, ...(required ? { required: true } : {}) })
const num = (description: string, required = false) => ({ type: 'number', description, ...(required ? { required: true } : {}) })
const bool = (description: string) => ({ type: 'boolean', description })
const cellMap = (description: string) => ({ type: 'object', additionalProperties: true, description, required: true })
const stringList = (description: string) => ({ type: 'array', items: { type: 'string' }, description, required: true })
const borderEdgeSchema = {
  type: 'object',
  additionalProperties: false,
  properties: {
    style: { type: 'string', enum: ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'], description: 'Border line style.' },
    color: text('Border color hex.'),
  },
}

const opSchema = (op: string, properties: Record<string, unknown>): any => ({
  type: 'object',
  additionalProperties: false,
  properties: {
    op: { type: 'string', enum: [op], required: true, description: `Operation: ${op}` },
    ...properties,
  },
})

const rangeSchema = text('Range with sheet, e.g. "Sheet1!A1:B10".', true)

export const excelOperationSchema = {
  oneOf: [
    opSchema('set', {
      cells: cellMap('Map of cell id (e.g. "Sheet1!A1") to content: numbers/dates/booleans are typed, strings starting with "=" become formulas.'),
    }),
    opSchema('fill', {
      source: text('Source cell id, e.g. "Sheet1!D2".', true),
      target: rangeSchema,
    }),
    opSchema('fillSeries', {
      start: text('Start cell id holding a number or date; must be the top-left cell of the target range.', true),
      target: rangeSchema,
      step: num('Step (default 1 for numbers, 1 day for dates).'),
    }),
    opSchema('insertRows', {
      sheet: text('Sheet name.', true),
      row: num('Insertion row (1-based).', true),
      count: num('Number of rows.', true),
    }),
    opSchema('deleteRows', {
      sheet: text('Sheet name.', true),
      row: num('First deleted row (1-based).', true),
      count: num('Number of rows.', true),
    }),
    opSchema('insertColumns', {
      sheet: text('Sheet name.', true),
      column: text('Insertion column letter, e.g. "B".', true),
      count: num('Number of columns.', true),
    }),
    opSchema('deleteColumns', {
      sheet: text('Sheet name.', true),
      column: text('First deleted column letter.', true),
      count: num('Number of columns.', true),
    }),
    opSchema('copyRange', {
      source: rangeSchema,
      target: text('Destination top-left cell id, e.g. "Sheet1!E2".', true),
      move: bool('Set true to clear the source range after copying.'),
    }),
    opSchema('sortRange', {
      range: rangeSchema,
      keys: {
        type: 'array',
        required: true,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            column: text('Column letter inside the range, e.g. "B".', true),
            direction: { type: 'string', enum: ['asc', 'desc'], description: 'Sort direction (default asc).' },
          },
        },
      },
      headerRows: num('Number of header rows to keep in place (default 0).'),
    }),
    opSchema('report', {
      source: text('Source data range including the header row, e.g. "订单!A1:F7".', true),
      groupColumn: text('Column letter to group and summarize by.', true),
      metrics: {
        type: 'array',
        required: true,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            column: text('Metric column letter.', true),
            function: { type: 'string', enum: ['sum', 'average', 'count', 'counta', 'max', 'min'], required: true, description: 'Aggregation function for the summary and subtotals.' },
          },
        },
      },
      sort: bool('Sort by the group column first (default true).'),
      subtotal: bool('Insert SUBTOTAL summary rows per group plus a grand total (default true).'),
      autoFilter: bool('Add filters to the header (default true).'),
      headerStyle: bool('Bold header with a light fill (default true).'),
      freezeHeader: bool('Freeze the header row (default true).'),
      numberFormat: text('Optional number format for metric cells, e.g. "#,##0.00".'),
      outputSheet: text('Summary sheet name (default "<source>汇总").'),
    }),
    opSchema('style', {
      range: rangeSchema,
      style: {
        type: 'object',
        required: true,
        additionalProperties: false,
        properties: {
          bold: bool('Bold text.'),
          italic: bool('Italic text.'),
          underline: bool('Underline text.'),
          fontSize: num('Font size in points.'),
          fontName: text('Font name, e.g. "Arial" or "微软雅黑".'),
          fontColor: text('Font color as 6-digit hex, e.g. "FF0000".'),
          fill: text('Background fill color as 6-digit hex, e.g. "D9D9D9".'),
          numberFormat: text('Excel number format, e.g. "#,##0.00" or "0.00%".'),
          hAlign: { type: 'string', enum: ['left', 'center', 'right'], description: 'Horizontal alignment.' },
          vAlign: { type: 'string', enum: ['top', 'middle', 'bottom'], description: 'Vertical alignment.' },
          wrapText: bool('Wrap text.'),
          border: {
            type: 'object',
            additionalProperties: false,
            properties: {
              top: borderEdgeSchema,
              bottom: borderEdgeSchema,
              left: borderEdgeSchema,
              right: borderEdgeSchema,
            },
          },
        },
      },
    }),
    opSchema('dataValidation', {
      range: rangeSchema,
      type: { type: 'string', enum: ['list', 'whole', 'decimal', 'date', 'textLength', 'custom'], required: true, description: 'Validation type.' },
      operator: { type: 'string', enum: ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'], description: 'Comparison operator for non-list types.' },
      formula1: text('For list: comma-separated items ("高,中,低") or a range. For others: lower bound or value.'),
      formula2: text('Upper bound for between / notBetween.'),
      allowBlank: bool('Allow empty cells.'),
      showInputMessage: bool('Show an input prompt.'),
      prompt: text('Input prompt text.'),
      showErrorMessage: bool('Reject invalid input.'),
      errorStyle: { type: 'string', enum: ['stop', 'warning', 'information'], description: 'Error alert style.' },
      error: text('Error message text.'),
      errorTitle: text('Error alert title.'),
    }),
    opSchema('conditionalFormatting', {
      range: rangeSchema,
      rules: {
        type: 'array',
        required: true,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            type: {
              type: 'string',
              enum: ['cellIs', 'expression', 'containsText', 'dataBar', 'colorScale', 'iconSet', 'top10'],
              required: true,
              description: 'Rule type: cellIs / expression / containsText / dataBar / colorScale / iconSet / top10.',
            },
            operator: text('cellIs operator, e.g. "greaterThan", "lessThan", "equal", "between".'),
            formula: text('Threshold value (number as text) or expression formula.'),
            formula2: text('Upper bound for between.'),
            text: text('Text to match for containsText.'),
            color: text('Bar color hex for dataBar.'),
            minColor: text('Low color hex for colorScale.'),
            midColor: text('Mid color hex for colorScale.'),
            maxColor: text('High color hex for colorScale.'),
            iconSet: text('Icon set name for iconSet, e.g. "3Arrows".'),
            rank: num('Rank for top10 (default 10).'),
            percent: bool('top10 as top percent.'),
            bottom: bool('top10 as bottom 10.'),
            style: {
              type: 'object',
              additionalProperties: false,
              properties: {
                bold: bool('Bold text.'),
                fontColor: text('Font color hex.'),
                fill: text('Fill color hex.'),
              },
            },
          },
        },
      },
    }),
    opSchema('autoFilter', { range: rangeSchema }),
    opSchema('subtotal', {
      sheet: text('Sheet name.', true),
      range: rangeSchema,
      groupColumn: text('Column letter to group by; the data should be sorted by it.', true),
      summaryColumns: {
        type: 'array',
        required: true,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            column: text('Column letter to summarize.', true),
            function: { type: 'string', enum: ['sum', 'average', 'count', 'max', 'min'], required: true, description: 'Summary function.' },
          },
        },
      },
      addGrandTotal: bool('Add a grand total row (default true).'),
    }),
    opSchema('aggregateReport', {
      source: text('Source data range including the header row, e.g. "Sheet1!A1:C50".', true),
      groupColumn: text('Column letter to group by (e.g. "A").', true),
      metrics: {
        type: 'array',
        required: true,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            column: text('Column letter of the metric.', true),
            function: { type: 'string', enum: ['sum', 'average', 'count', 'counta', 'max', 'min'], required: true, description: 'Aggregation; the report uses live SUMIFS/AVERAGEIFS/COUNTIFS/MAXIFS/MINIFS formulas so it stays dynamic.' },
          },
        },
      },
      outputSheet: text('Output sheet name (default "<source>汇总").'),
    }),
    opSchema('filterToRange', {
      source: text('Source data range including the header row.', true),
      criteria: {
        type: 'array',
        required: true,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            column: text('Column letter to test.', true),
            operator: { type: 'string', enum: ['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'contains'], required: true, description: 'Comparison operator.' },
            value: { oneOf: [{ type: 'string' }, { type: 'number' }], description: 'Comparison value.' },
          },
        },
      },
      target: text('Destination top-left cell, e.g. "Sheet1!D1".', true),
      matchAll: bool('AND across criteria (default true); false means OR.'),
    }),
    opSchema('protectSheet', {
      sheet: text('Sheet name.', true),
      password: text('Optional password; empty means protected without a password.'),
      options: {
        type: 'object',
        additionalProperties: false,
        properties: {
          selectLockedCells: bool('Allow selecting locked cells (default true).'),
          selectUnlockedCells: bool('Allow selecting unlocked cells (default true).'),
          formatCells: bool('Allow formatting cells.'),
          formatColumns: bool('Allow formatting columns.'),
          formatRows: bool('Allow formatting rows.'),
          insertColumns: bool('Allow inserting columns.'),
          insertRows: bool('Allow inserting rows.'),
          deleteColumns: bool('Allow deleting columns.'),
          deleteRows: bool('Allow deleting rows.'),
          sort: bool('Allow sorting.'),
          autoFilter: bool('Allow using filters.'),
        },
      },
    }),
    opSchema('unprotectSheet', {
      sheet: text('Sheet name.', true),
      password: text('Password used when protecting (if any).'),
    }),
    opSchema('mailMerge', {
      template: text('Template range with {Placeholder} tokens, e.g. "Sheet1!A1:B1".', true),
      data: text('Data range with a header row matching the placeholders, e.g. "Sheet2!A1:C10".', true),
      outputSheet: text('Output sheet name (default "<template>合并").'),
    }),
    opSchema('pageSetup', {
      sheet: text('Sheet name.', true),
      printArea: text('Print area range without sheet, e.g. "A1:F40".'),
      orientation: { type: 'string', enum: ['portrait', 'landscape'], description: 'Page orientation.' },
      fitToPage: bool('Fit to page width/height.'),
      fitToWidth: num('Pages wide for fit-to-page.'),
      fitToHeight: num('Pages tall for fit-to-page.'),
      margins: {
        type: 'object',
        additionalProperties: false,
        properties: {
          top: num('Top margin in inches.'),
          right: num('Right margin in inches.'),
          bottom: num('Bottom margin in inches.'),
          left: num('Left margin in inches.'),
          header: num('Header margin in inches.'),
          footer: num('Footer margin in inches.'),
        },
      },
      centerHorizontally: bool('Center horizontally on the page.'),
      centerVertically: bool('Center vertically on the page.'),
    }),
    opSchema('definedName', {
      name: text('Name without spaces, e.g. "SalesRange".', true),
      ref: text('Absolute reference, e.g. "Sheet1!$A$1:$D$50".', true),
    }),
    opSchema('addTable', {
      name: text('Unique table name.', true),
      range: rangeSchema,
      headerRow: bool('Use the first row as headers (default true).'),
      totalsRow: bool('Add a totals row (default false).'),
      showRowStripes: bool('Zebra stripes (default true).'),
      showColumnStripes: bool('Column stripes (default false).'),
    }),
    opSchema('setColumnWidth', {
      sheet: text('Sheet name.', true),
      column: text('Column letter.', true),
      width: num('Width in characters.', true),
    }),
    opSchema('setRowHeight', {
      sheet: text('Sheet name.', true),
      row: num('Row number.', true),
      height: num('Height in points.', true),
    }),
    opSchema('freezePanes', {
      sheet: text('Sheet name.', true),
      row: num('First scrollable row (1-based; 2 freezes row 1).', true),
      column: text('First scrollable column letter (1-based; "B" freezes column A).', true),
    }),
    opSchema('findReplace', {
      find: text('Text to find.', true),
      replace: text('Replacement text.', true),
      sheet: text('Restrict to one sheet (optional).'),
      matchCase: bool('Case-sensitive match (default false).'),
    }),
    opSchema('addSheet', { name: text('New sheet name.', true) }),
    opSchema('renameSheet', {
      oldName: text('Current sheet name.', true),
      newName: text('New sheet name; formulas referencing the old name are updated.', true),
    }),
    opSchema('deleteSheet', { name: text('Sheet name to delete.', true) }),
    opSchema('duplicateSheet', {
      name: text('Source sheet name.', true),
      newName: text('New sheet name.', true),
    }),
    opSchema('hideSheet', {
      name: text('Sheet name.', true),
      hidden: bool('Hide (default true) or show (false).'),
    }),
    opSchema('setTabColor', {
      name: text('Sheet name.', true),
      color: text('Tab color as 6-digit hex.', true),
    }),
    opSchema('clear', { cells: stringList('Cell ids to clear, e.g. ["Sheet1!A1"].') }),
    opSchema('merge', { range: rangeSchema }),
    opSchema('unmerge', { range: rangeSchema }),
  ],
} as const
