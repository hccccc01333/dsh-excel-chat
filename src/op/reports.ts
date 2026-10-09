/**
 * Report-shaped operations: subtotals, the one-shot report, role presets and the
 * live-formula aggregate table.
 *
 * These share a shape — group by a column, aggregate others, write a summary sheet
 * whose cells are SUMIFS rather than values — so they are read together or not at
 * all. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs'
import { columnToNumber, numberToColumn } from '../formula.ts'
import { absoluteColumnRef, findSheet, parseRange } from './core.ts'
import { applyConditionalFormatting, applyStyle } from './formatting.ts'
import { applyFilterToRange } from './filters.ts'
import { sortRange } from './sorting.ts'
import { shiftWorkbookRows } from './structure.ts'
import { t } from '../i18n.ts'
import { cellContent } from '../workbook.ts'
import type { ExcelOperation } from '../operation-types.ts'
const SUBTOTAL_CODES: Record<string, number> = {
  sum: 9,
  average: 1,
  count: 2,
  max: 4,
  min: 5,
}

export function applySubtotal(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'subtotal' }>,
): number {
  const parsed = parseRange(workbook, options.range)
  const groupCol = columnToNumber(options.groupColumn)
  if (groupCol < parsed.startCol || groupCol > parsed.endCol) {
    throw new Error(`subtotal group column outside range: ${options.groupColumn}`)
  }
  for (const summary of options.summaryColumns) {
    const col = columnToNumber(summary.column)
    if (col < parsed.startCol || col > parsed.endCol) {
      throw new Error(`subtotal summary column outside range: ${summary.column}`)
    }
    if (!SUBTOTAL_CODES[summary.function]) throw new Error(`unsupported subtotal function: ${summary.function}`)
  }
  const sheet = parsed.sheet
  const header = parsed.startRow
  const firstData = parsed.startRow + 1
  const lastData = parsed.endRow

  interface SubtotalGroup {
    value: string
    startRow: number
    endRow: number
  }
  const groups: SubtotalGroup[] = []
  let current: SubtotalGroup | null = null
  for (let row = firstData; row <= lastData; row++) {
    const raw = sheet.getCell(`${numberToColumn(groupCol)}${row}`).value
    const key = raw === null || raw === undefined ? '' : String(raw)
    if (!current || current.value !== key) {
      current = { value: key, startRow: row, endRow: row }
      groups.push(current)
    } else {
      current.endRow = row
    }
  }

  for (let groupIndex = 0; groupIndex < groups.length; groupIndex++) {
    const group = groups[groupIndex]!
    const finalStartRow = group.startRow + groupIndex
    const finalEndRow = group.endRow + groupIndex
    const insertRow = finalEndRow + 1
    sheet.spliceRows(insertRow, 0, [])
    shiftWorkbookRows(workbook, sheet.name, insertRow, 1)
    const label = sheet.getCell(`${numberToColumn(groupCol)}${insertRow}`)
    label.value = `${group.value} 汇总`
    label.font = { bold: true }
    for (const summary of options.summaryColumns) {
      const col = columnToNumber(summary.column)
      const cell = sheet.getCell(`${numberToColumn(col)}${insertRow}`)
      cell.value = {
        formula: `SUBTOTAL(${SUBTOTAL_CODES[summary.function]},${numberToColumn(col)}${finalStartRow}:${numberToColumn(col)}${finalEndRow})`,
      }
      cell.font = { bold: true }
    }
  }

  if (options.addGrandTotal ?? true) {
    const totalRow = parsed.endRow + groups.length + 1
    sheet.spliceRows(totalRow, 0, [])
    shiftWorkbookRows(workbook, sheet.name, totalRow, 1)
    const label = sheet.getCell(`${numberToColumn(groupCol)}${totalRow}`)
    label.value = '总计'
    label.font = { bold: true }
    for (const summary of options.summaryColumns) {
      const col = columnToNumber(summary.column)
      const cell = sheet.getCell(`${numberToColumn(col)}${totalRow}`)
      cell.value = {
        formula: `SUBTOTAL(${SUBTOTAL_CODES[summary.function]},${numberToColumn(col)}${firstData}:${numberToColumn(col)}${lastData + groups.length})`,
      }
      cell.font = { bold: true }
    }
  }
  void header
  return groups.length + (options.addGrandTotal ?? true ? 1 : 0)
}

/**
 * One-shot report template: sort, subtotals, a dynamic SUMIFS summary sheet,
 * auto filter, header style, frozen header, and optional number format.
 * Ordering matters: subtotals run before the summary so its SUMIFS ranges
 * already cover the final data block (subtotal rows do not match group keys).
 */
export function applyReport(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'report' }>,
): void {
  const parsed = parseRange(workbook, options.source)
  const sheet = parsed.sheet
  const groupCol = columnToNumber(options.groupColumn)
  if (groupCol < parsed.startCol || groupCol > parsed.endCol) {
    throw new Error(`report group column outside range: ${options.groupColumn}`)
  }
  if (options.sort ?? true) {
    sortRange(workbook, options.source, [{ column: options.groupColumn }], 1)
  }
  let finalEndRow = parsed.endRow
  if (options.subtotal ?? true) {
    const subtotalMetrics = options.metrics.map((metric) => ({
      column: metric.column,
      function: metric.function === 'counta' ? 'count' : metric.function,
    })) as Array<{ column: string; function: 'sum' | 'average' | 'count' | 'max' | 'min' }>
    const inserted = applySubtotal(workbook, {
      op: 'subtotal',
      sheet: sheet.name,
      range: options.source,
      groupColumn: options.groupColumn,
      summaryColumns: subtotalMetrics,
      addGrandTotal: true,
    })
    finalEndRow = parsed.endRow + inserted
  }
  const summarySheet = options.outputSheet ?? `${sheet.name}-汇总`
  applyAggregateReport(workbook, {
    op: 'aggregateReport',
    source: `${sheet.name}!${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${finalEndRow}`,
    groupColumn: options.groupColumn,
    metrics: options.metrics,
    outputSheet: summarySheet,
  })
  if (options.autoFilter ?? true) {
    sheet.autoFilter = {
      from: { row: parsed.startRow, column: parsed.startCol },
      to: { row: finalEndRow, column: parsed.endCol },
    }
  }
  if (options.headerStyle ?? true) {
    applyStyle(workbook, `${sheet.name}!${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${parsed.startRow}`, {
      bold: true,
      fill: 'D9D9D9',
    })
  }
  if (options.freezeHeader ?? true) {
    sheet.views = [{
      state: 'frozen',
      xSplit: Math.max(0, parsed.startCol - 1),
      ySplit: Math.max(0, parsed.startRow),
      topLeftCell: `${numberToColumn(parsed.startCol)}${parsed.startRow + 1}`,
    }]
  }
  if (options.numberFormat) {
    for (const metric of options.metrics) {
      const col = columnToNumber(metric.column)
      for (let row = parsed.startRow; row <= finalEndRow; row++) {
        sheet.getCell(`${numberToColumn(col)}${row}`).numFmt = options.numberFormat
      }
    }
    const summary = findSheet(workbook, summarySheet)
    if (summary) {
      options.metrics.forEach((metric, index) => {
        const col = numberToColumn(2 + index)
        for (let row = 1; row <= summary.rowCount; row++) {
          summary.getCell(`${col}${row}`).numFmt = options.numberFormat
        }
      })
    }
  }
}

const ROLE_LABELS: Record<string, string> = {
  ops: '运营报表',
  product: '产品分析',
  data: '数据分析',
}

/**
 * Role-based one-shot preset: 运营 gets a report with data bars, 产品 and 数分
 * get a report with color scales, and 数分 additionally writes a filtered copy.
 */
export function applyPreset(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'preset' }>,
): void {
  const parsed = parseRange(workbook, options.source)
  const sheet = parsed.sheet
  const summarySheet = `${sheet.name}-${ROLE_LABELS[options.role]}`
  if (options.filter) {
    const filterSheetName = `${sheet.name}-筛选`
    if (!findSheet(workbook, filterSheetName)) workbook.addWorksheet(filterSheetName)
    applyFilterToRange(workbook, {
      op: 'filterToRange',
      source: options.source,
      criteria: [options.filter],
      target: `${filterSheetName}!A1`,
    })
  }
  applyReport(workbook, {
    op: 'report',
    source: options.source,
    groupColumn: options.groupColumn,
    metrics: options.metrics,
    numberFormat: '#,##0.00',
    outputSheet: summarySheet,
  })
  for (const metric of options.metrics) {
    const col = numberToColumn(columnToNumber(metric.column))
    const range = `${sheet.name}!${col}${parsed.startRow}:${col}${sheet.rowCount}`
    if (options.role === 'ops') {
      applyConditionalFormatting(workbook, range, [{ type: 'dataBar', color: '63BE7B' }])
    } else {
      applyConditionalFormatting(workbook, range, [{
        type: 'colorScale',
        minColor: 'F8696B',
        midColor: 'FFEB84',
        maxColor: '63BE7B',
      }])
    }
  }
}

const REPORT_FUNCTIONS: Record<string, string> = {
  sum: 'SUMIFS',
  average: 'AVERAGEIFS',
  count: 'COUNTIFS',
  counta: 'COUNTIFS',
  max: 'MAXIFS',
  min: 'MINIFS',
}

/**
 * True when a row carries SUBTOTAL formulas. That is the marker `applySubtotal`
 * writes, and the only subtotal marker that does not depend on the label's
 * language — the labels themselves are workbook data and stay untranslated.
 */
function isSubtotalRow(sheet: ExcelJS.Worksheet, row: number, metricColumns: readonly string[]): boolean {
  return metricColumns.some((column) => {
    const value = sheet.getCell(`${numberToColumn(columnToNumber(column))}${row}`).value
    const formula = value !== null && typeof value === 'object' ? (value as { formula?: string }).formula : undefined
    return typeof formula === 'string' && formula.trimStart().toUpperCase().startsWith('SUBTOTAL(')
  })
}

export function applyAggregateReport(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'aggregateReport' }>,
): void {
  const parsed = parseRange(workbook, options.source)
  const groupCol = columnToNumber(options.groupColumn)
  const sourceSheet = parsed.sheet.name
  const firstData = parsed.startRow + 1
  const lastData = parsed.endRow
  const groupRange = absoluteColumnRef(sourceSheet, numberToColumn(groupCol), firstData, lastData)

  const groupValues: string[] = []
  const seen = new Set<string>()
  for (let row = firstData; row <= lastData; row++) {
    // Skip subtotal rows. `report` inserts them into the source before building
    // this summary (so the SUMIFS ranges below cover the final, sorted block),
    // which means the block handed to us contains derived rows as well as data.
    // Their label is "华东 汇总" / "总计" — a *different* string from every real
    // group key — so without this check each one becomes a phantom group, and the
    // grand-total row then sums the real groups twice over. Detected by the
    // SUBTOTAL formula rather than by the label, because the label is data and
    // therefore language-dependent.
    if (isSubtotalRow(parsed.sheet, row, options.metrics.map((metric) => metric.column))) continue
    const raw = parsed.sheet.getCell(`${numberToColumn(groupCol)}${row}`).value
    const key = raw === null || raw === undefined ? '' : String(raw)
    if (!seen.has(key)) {
      seen.add(key)
      groupValues.push(key)
    }
  }

  const outputSheetName = options.outputSheet ?? `${sourceSheet}-汇总`
  let output = findSheet(workbook, outputSheetName)
  if (!output) output = workbook.addWorksheet(outputSheetName)
  const groupHeader = String(parsed.sheet.getCell(`${numberToColumn(groupCol)}${parsed.startRow}`).value ?? options.groupColumn)
  output.getCell('A1').value = groupHeader
  output.getCell('A1').font = { bold: true }
  const metricLabels: Record<string, string> = {
    sum: '合计',
    average: '平均',
    count: '计数',
    counta: '非空计数',
    max: '最大',
    min: '最小',
  }
  options.metrics.forEach((metric, index) => {
    const metricCol = columnToNumber(metric.column)
    const header = String(parsed.sheet.getCell(`${numberToColumn(metricCol)}${parsed.startRow}`).value ?? metric.column)
    const cell = output.getCell(`${numberToColumn(2 + index)}1`)
    cell.value = `${header} ${metricLabels[metric.function]}`
    cell.font = { bold: true }
    void metricCol
  })

  for (let index = 0; index < groupValues.length; index++) {
    const row = 2 + index
    const groupCell = output.getCell(`A${row}`)
    groupCell.value = groupValues[index]
    options.metrics.forEach((metric, metricIndex) => {
      const metricCol = columnToNumber(metric.column)
      const metricRange = absoluteColumnRef(sourceSheet, numberToColumn(metricCol), firstData, lastData)
      const fn = REPORT_FUNCTIONS[metric.function]!
      const criteria = `A${row}`
      output.getCell(`${numberToColumn(2 + metricIndex)}${row}`).value = {
        formula: `${fn}(${metricRange},${groupRange},${criteria})`,
      }
    })
  }
  const lastGroupRow = 1 + groupValues.length
  options.metrics.forEach((metric, metricIndex) => {
    const col = numberToColumn(2 + metricIndex)
    output.getCell(`${col}${lastGroupRow + 1}`).value = {
      formula: `SUM(${col}2:${col}${lastGroupRow})`,
    }
    output.getCell(`${col}${lastGroupRow + 1}`).font = { bold: true }
  })
  output.getCell(`A${lastGroupRow + 1}`).value = '总计'
  output.getCell(`A${lastGroupRow + 1}`).font = { bold: true }
}

