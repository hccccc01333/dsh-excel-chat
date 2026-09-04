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
    style: {
      type: 'string',
      enum: [
        'thin', 'medium', 'thick', 'dashed', 'dotted', 'double',
        'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'slantDashDot',
      ],
      description: 'Border line style.',
    },
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
      valuesOnly: bool('Paste values only: formulas contribute their last cached result (default false).'),
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
    opSchema('preset', {
      role: { type: 'string', enum: ['ops', 'product', 'data'], required: true, description: '岗位：ops=运营 / product=产品 / data=数分。' },
      source: text('Source data range including the header row, e.g. "订单!A1:F7".', true),
      groupColumn: text('Column letter to group by.', true),
      metrics: {
        type: 'array',
        required: true,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            column: text('Metric column letter.', true),
            function: { type: 'string', enum: ['sum', 'average', 'count', 'counta', 'max', 'min'], required: true, description: 'Aggregation function.' },
          },
        },
      },
      filter: {
        type: 'object',
        additionalProperties: false,
        properties: {
          column: text('Column letter to filter by (数分 preset writes a filtered copy).', true),
          operator: { type: 'string', enum: ['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'contains'], required: true, description: 'Comparison operator.' },
          value: { oneOf: [{ type: 'string' }, { type: 'number' }], required: true, description: 'Comparison value.' },
        },
      },
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
          strikeThrough: bool('Strikethrough text.'),
          fontSize: num('Font size in points.'),
          fontName: text('Font name, e.g. "Arial" or "微软雅黑".'),
          fontColor: text('Font color as 6-digit hex, e.g. "FF0000".'),
          fill: text('Background fill color as 6-digit hex, e.g. "D9D9D9".'),
          numberFormat: text('Excel number format, e.g. "#,##0.00" or "0.00%".'),
          hAlign: { type: 'string', enum: ['left', 'center', 'right'], description: 'Horizontal alignment.' },
          vAlign: { type: 'string', enum: ['top', 'middle', 'bottom'], description: 'Vertical alignment.' },
          wrapText: bool('Wrap text.'),
          textRotation: num('Text rotation in degrees (0-180, 255 = vertical).'),
          shrinkToFit: bool('Shrink text to fit the cell.'),
          indent: num('Indent level 0-15 for left/right alignment.'),
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
              enum: [
                'cellIs', 'expression', 'containsText', 'notContainsText', 'blanks', 'noBlanks',
                'errors', 'noErrors', 'duplicateValues', 'uniqueValues', 'aboveAverage',
                'belowAverage', 'timePeriod', 'dataBar', 'colorScale', 'iconSet', 'top10',
              ],
              required: true,
              description: 'Rule type: cellIs / expression / containsText / notContainsText / blanks / noBlanks / errors / noErrors / duplicateValues / uniqueValues / aboveAverage / belowAverage / timePeriod / dataBar / colorScale / iconSet / top10.',
            },
            operator: text('cellIs operator, e.g. "greaterThan", "lessThan", "equal", "between".'),
            formula: text('Threshold value (number as text) or expression formula.'),
            formula2: text('Upper bound for between.'),
            text: text('Text to match for containsText.'),
            timePeriod: { type: 'string', enum: ['today', 'yesterday', 'tomorrow', 'last7Days', 'thisMonth', 'lastMonth', 'nextMonth', 'thisWeek', 'lastWeek', 'nextWeek'], description: 'Date period for timePeriod.' },
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
    opSchema('importCsv', {
      file: text('Absolute path to the CSV file to import.', true),
      sheet: text('Target sheet name (default a new "CSV" sheet).'),
      delimiter: text('Field delimiter (default ",").'),
      firstRowHeaders: bool('Whether the first row is a header row (metadata only).'),
    }),
    opSchema('exportCsv', {
      file: text('Absolute path to write the CSV file.', true),
      sheet: text('Source sheet name (default first sheet).'),
      range: text('Optional range to export, e.g. "A1:F7".'),
      delimiter: text('Field delimiter (default ",").'),
      guardFormulas: bool('Prefix literal values starting with = + - @ to prevent formula injection (default true).'),
    }),
    opSchema('clear', { cells: stringList('Cell ids to clear, e.g. ["Sheet1!A1"].') }),
    opSchema('merge', { range: rangeSchema }),
    opSchema('unmerge', { range: rangeSchema }),
    opSchema('dedupeRows', {
      sheet: text('Sheet name.', true),
      columns: {
        type: 'array',
        items: { type: 'string' },
        description: 'Column letters that define a duplicate (default: all columns).',
      },
      keep: { type: 'string', enum: ['first', 'last'], description: 'Which occurrence to keep (default first).' },
    }),
    opSchema('fillMissing', {
      range: rangeSchema,
      mode: { type: 'string', enum: ['value', 'forward', 'left'], required: true, description: 'Fill mode: value (fixed value), forward (copy nearest non-empty cell above), left (copy nearest non-empty cell to the left).' },
      value: { oneOf: [{ type: 'string' }, { type: 'number' }], description: 'Fill value, required when mode is "value".' },
    }),
    opSchema('removeEmptyRows', { range: rangeSchema }),
    opSchema('removeEmptyColumns', { range: rangeSchema }),
    opSchema('trimText', { range: rangeSchema }),
    opSchema('changeCase', {
      range: rangeSchema,
      case: { type: 'string', enum: ['upper', 'lower', 'proper'], required: true, description: 'Case conversion: upper / lower / proper (capitalize each word).' },
    }),
    opSchema('normalizeText', { range: rangeSchema }),
    opSchema('splitColumn', {
      sheet: text('Sheet name.', true),
      column: text('Source column letter, e.g. "A".', true),
      delimiter: text('Delimiter to split on, e.g. "-" or " ".', true),
      startRow: num('First data row (1-based).', true),
      endRow: num('Last data row (default: last used row).'),
    }),
    opSchema('highlightRows', {
      sheet: text('Sheet name.', true),
      range: rangeSchema,
      criteria: {
        type: 'array',
        required: true,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            column: text('Column letter to test, e.g. "A".', true),
            operator: { type: 'string', enum: ['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'contains'], required: true, description: 'Comparison operator.' },
            value: { oneOf: [{ type: 'string' }, { type: 'number' }], required: true, description: 'Comparison value.' },
          },
        },
        description: 'Rows matching ALL criteria are highlighted.',
      },
      style: {
        type: 'object',
        additionalProperties: true,
        description: 'Optional highlight style (default yellow fill).',
      },
    }),
    opSchema('fuzzyMatch', {
      source: rangeSchema,
      sourceKey: text('Source key column letter, e.g. "A".', true),
      target: rangeSchema,
      targetKey: text('Target key column letter.', true),
      valueColumn: text('Target column whose value is copied back on a match.', true),
      outputColumn: text('Source column letter where the matched value is written.', true),
      threshold: num('Minimum similarity in [0,1] to accept a match (default 0.6).'),
      scoreColumn: text('Optional source column letter for the match score.'),
    }),
    opSchema('hideRows', {
      sheet: text('Sheet name.', true),
      from: num('First row to hide/show (1-based).', true),
      to: num('Last row to hide/show (1-based).', true),
      hidden: bool('Hide (default true) or show (false).'),
    }),
    opSchema('hideColumns', {
      sheet: text('Sheet name.', true),
      columns: { type: 'array', items: { type: 'string' }, required: true, description: 'Column letters to hide/show, e.g. ["B","D"] or ["C"].' },
      hidden: bool('Hide (default true) or show (false).'),
    }),
    opSchema('groupRows', {
      sheet: text('Sheet name.', true),
      start: num('First row of the group (1-based).', true),
      end: num('Last row of the group (inclusive).', true),
      level: num('Outline level 1-7; 0 ungroups (default 1).'),
      collapse: bool('Collapse the group after grouping.'),
    }),
    opSchema('groupColumns', {
      sheet: text('Sheet name.', true),
      from: text('First grouped column letter, e.g. "B".', true),
      to: text('Last grouped column letter (inclusive).', true),
      level: num('Outline level 1-7; 0 ungroups (default 1).'),
      collapse: bool('Collapse the group after grouping.'),
    }),
    opSchema('autoFitColumnWidths', {
      sheet: text('Sheet name.', true),
      columns: { type: 'array', items: { type: 'string' }, description: 'Column letters to fit (default all used columns).' },
      minWidth: num('Minimum width in characters (default 8).'),
      maxWidth: num('Maximum width in characters (default 60); CJK characters count double.'),
    }),
    opSchema('unfreezePanes', { sheet: text('Sheet name.', true) }),
    opSchema('transpose', {
      source: rangeSchema,
      target: text('Destination top-left cell id, e.g. "Sheet2!A1".', true),
    }),
    opSchema('clearRange', {
      range: rangeSchema,
      mode: { type: 'string', enum: ['contents', 'formats', 'all'], description: 'contents = values/formulas only (styles kept), formats = styles/number formats only, all = everything (default contents).' },
    }),
    opSchema('joinSheets', {
      source: rangeSchema,
      sourceKey: text('Source key column letter, e.g. "A".', true),
      lookup: rangeSchema,
      lookupKey: text('Lookup table key column letter.', true),
      valueColumns: stringList('Lookup table columns whose values are copied back, e.g. ["B","D"].'),
      outputColumns: stringList('Source columns to write each matched value into (same order and count as valueColumns), e.g. ["F","G"].'),
      missValue: { oneOf: [{ type: 'string' }, { type: 'number' }], description: 'Optional fallback written when no lookup row matches (default: leave blank).' },
    }),
    opSchema('crosstab', {
      source: rangeSchema,
      rowColumn: text('Column letter whose distinct values become the output rows.', true),
      columnColumn: text('Column letter whose distinct values become the output columns.', true),
      metric: {
        type: 'object',
        required: true,
        additionalProperties: false,
        properties: {
          column: text('Metric column letter (required for sum / average / max / min).'),
          function: { type: 'string', enum: ['sum', 'average', 'count', 'counta', 'max', 'min'], required: true, description: 'Aggregation applied per row x column cell via live SUMIFS/COUNTIFS formulas.' },
        },
      },
      outputSheet: text('Output sheet name (default "<source>交叉表").'),
      totals: bool('Add grand total row and column (default true; only for sum / count / counta).'),
    }),
    opSchema('setHyperlink', {
      cell: text('Cell id, e.g. "Sheet1!A1".', true),
      url: text('External URL, e.g. "https://example.com".'),
      location: text('Internal link target without #, e.g. "Sheet2!A1" or "明细!A1".'),
      text: text('Display text (default url or location).'),
    }),
    opSchema('printTitles', {
      sheet: text('Sheet name.', true),
      rows: text('Rows repeated at the top of every printed page, e.g. "1:1" or "1:3".'),
      columns: text('Columns repeated on the left of every printed page, e.g. "A:A" or "A:B".'),
    }),
    opSchema('copyStyle', {
      source: text('Source cell id whose format is copied, e.g. "Sheet1!A1".', true),
      target: rangeSchema,
    }),
    opSchema('freezeFormulas', { range: rangeSchema }),
    opSchema('uniqueValues', {
      source: text('Source range; the FIRST column of it supplies the values, e.g. "Sheet1!A1:A50".', true),
      target: text('Destination top-left cell id, e.g. "Sheet2!A1".', true),
      includeHeader: bool('Write the source header cell as the first output row (default false).'),
    }),
    opSchema('unmergeAll', { sheet: text('Sheet name.', true) }),
    opSchema('setZoom', {
      sheet: text('Sheet name.', true),
      zoom: num('Zoom level 10-400 (percent).', true),
      normalZoom: num('Zoom used when returning to normal view (default zoom).'),
    }),
    opSchema('showGridLines', {
      sheet: text('Sheet name.', true),
      visible: { type: 'boolean', required: true, description: 'true to show gridlines, false to hide them.' },
    }),
    opSchema('headerFooter', {
      sheet: text('Sheet name.', true),
      oddHeader: text('Header text with &-codes, e.g. "&L公司&C报表&R日期" (L/C/R = left/center/right sections).'),
      oddFooter: text('Footer text with &-codes, e.g. "第 &P 页 / 共 &N 页" (&P page, &N total pages).'),
      evenHeader: text('Even-page header (needs differentOddEven).'),
      evenFooter: text('Even-page footer (needs differentOddEven).'),
      firstHeader: text('First-page header (needs differentFirst).'),
      firstFooter: text('First-page footer (needs differentFirst).'),
      differentOddEven: bool('Use different headers/footers on odd and even pages.'),
      differentFirst: bool('Use a different header/footer on the first page.'),
    }),
    opSchema('moveSheet', {
      name: text('Sheet name to move.', true),
      position: num('Target position counting from 1 (1 = first tab).', true),
    }),
    opSchema('setWorkbookProperties', {
      creator: text('Author name.'),
      lastModifiedBy: text('Last modified by name.'),
      title: text('Document title.'),
      subject: text('Document subject.'),
      description: text('Document description.'),
      keywords: text('Comma-separated keywords.'),
      recalcOnOpen: bool('Force Excel to recalculate all formulas when the file opens.'),
    }),
    opSchema('rankColumn', {
      range: rangeSchema,
      metricColumn: text('Column letter to rank by (inside the range).', true),
      outputColumn: text('Column letter where RANK formulas are written, e.g. "E".', true),
      descending: bool('Rank 1 = largest value (default true); false ranks smallest first.'),
      skipHeader: bool('Range includes a header row (default true); set false for data-only ranges.'),
    }),
    opSchema('rowPageBreaks', {
      sheet: text('Sheet name.', true),
      rows: { type: 'array', items: { type: 'number' }, required: true, description: '1-based row numbers; a manual page break is inserted above each row, e.g. [11, 21] prints rows 1-10 then 11-20.' },
    }),
    opSchema('clearPageBreaks', { sheet: text('Sheet name.', true) }),
    opSchema('addComment', {
      cell: text('Cell id for the comment, e.g. "Sheet1!B2".', true),
      text: text('Comment text (plain text).', true),
      author: text('Author name (default dsh-excel-chat).'),
      width: num('Comment box width in points (default 108).'),
      height: num('Comment box height in points (default 60).'),
    }),
    opSchema('addSparklines', {
      dataRange: text('Data range with one row per sparkline, e.g. "Sheet1!B2:F31".', true),
      locationRange: text('Location range with the same number of rows, e.g. "Sheet1!G2:G31"; sparkline i renders row i of dataRange.', true),
      type: { type: 'string', enum: ['line', 'column', 'stacked'], description: 'Sparkline style (default line).' },
      color: text('Series color hex (default green 375623).'),
      negativeColor: text('Color for negative values (default red D00000).'),
      markers: bool('Show markers with high/low highlighting.'),
      highColor: text('Color for the highest point (default orange FF7C00).'),
      lowColor: text('Color for the lowest point (default red D00000).'),
    }),
  ],
} as const
