/**
 * Build a native pivot table from a source range without Microsoft Excel.
 *
 * ExcelJS has no pivot model and drops pivot parts when it rewrites a workbook,
 * so the actual OOXML is emitted by xml-postprocess. This module computes the
 * aggregated cells (so the file is readable before Excel refreshes anything)
 * and describes the cache/table parts that need to be injected.
 */
import type ExcelJS from 'exceljs'
import { columnToNumber, numberToColumn } from '../formula.ts'
import { plainContent } from '../workbook.ts'
import type { PivotAggregate, PivotCacheField, PivotMeasure, PivotSpec } from '../xml-postprocess.ts'
import type { ExcelOperation } from '../operation-types.ts'
import { cellContentOf, findSheet, parseNumericText, parseRange } from './core.ts'

type PivotOperation = Extract<ExcelOperation, { op: 'pivot' }>
type MeasureInput = { column: number; name: string; aggregate: PivotAggregate }

function aggregate(values: number[], operation: PivotAggregate): number {
  if (values.length === 0) return 0
  const sum = values.reduce((total, value) => total + value, 0)
  switch (operation) {
    case 'average': return sum / values.length
    case 'count': return values.length
    case 'max': return values.reduce((best, value) => Math.max(best, value), values[0]!)
    case 'min': return values.reduce((best, value) => Math.min(best, value), values[0]!)
    case 'sum': return sum
  }
}

const labels: Record<PivotAggregate, string> = {
  sum: '求和', average: '平均值', count: '计数', max: '最大值', min: '最小值',
}

export function buildPivot(
  workbook: ExcelJS.Workbook,
  operation: PivotOperation,
  cacheId: number,
  partSuffix: number,
): PivotSpec {
  const parsed = parseRange(workbook, operation.source)
  const top = Math.min(parsed.startRow, parsed.endRow)
  const bottom = Math.max(parsed.startRow, parsed.endRow)
  const left = Math.min(parsed.startCol, parsed.endCol)
  const right = Math.max(parsed.startCol, parsed.endCol)
  const rowColumn = columnToNumber(operation.rowField)
  const columnColumn = operation.columnField ? columnToNumber(operation.columnField) : undefined
  const pageColumn = operation.pageField ? columnToNumber(operation.pageField) : undefined
  const measureInputs: MeasureInput[] = [
    { column: columnToNumber(operation.valueField), name: '', aggregate: operation.aggregate ?? 'sum' },
    ...(operation.valueFields ?? []).map((field) => ({
      column: columnToNumber(field.column), name: '', aggregate: field.aggregate ?? 'sum',
    })),
  ]
  for (const [label, column] of [
    ['rowField', rowColumn],
    ...(columnColumn === undefined ? [] : [['columnField', columnColumn] as [string, number]]),
    ...(pageColumn === undefined ? [] : [['pageField', pageColumn] as [string, number]]),
    ...measureInputs.map((measure, index) => [`valueFields[${index}]`, measure.column] as const),
  ] as Array<[string, number]>) {
    if (column < left || column > right) {
      throw new Error(`pivot ${label} ${numberToColumn(column)} is outside range ${operation.source}`)
    }
  }

  const indexes = new Map<number, number>()
  const fields: PivotCacheField[] = []
  for (let col = left; col <= right; col += 1) {
    indexes.set(col, fields.length)
    const name = plainContent(cellContentOf(parsed.sheet.getCell(top, col))) || numberToColumn(col)
    const numeric = measureInputs.some((measure) => measure.column === col)
    fields.push(numeric ? { name, numeric: true, min: 0, max: 0 } : { name, numeric: false, items: [] })
  }
  for (const measure of measureInputs) {
    measure.name = fields[indexes.get(measure.column)!]!.name
  }
  const rowFieldIndex = indexes.get(rowColumn)!
  const rowField = fields[rowFieldIndex]!
  const columnFieldIndex = columnColumn === undefined ? undefined : indexes.get(columnColumn)
  if (columnFieldIndex !== undefined && (columnFieldIndex === rowFieldIndex || measureInputs.some((measure) => indexes.get(measure.column) === columnFieldIndex))) {
    throw new Error('pivot columnField must differ from rowField and the value columns')
  }
  const columnField = columnFieldIndex === undefined ? undefined : fields[columnFieldIndex]
  const pageFieldIndex = pageColumn === undefined ? undefined : indexes.get(pageColumn)
  if (pageFieldIndex !== undefined && (pageFieldIndex === rowFieldIndex || pageFieldIndex === columnFieldIndex
    || measureInputs.some((measure) => indexes.get(measure.column) === pageFieldIndex))) {
    throw new Error('pivot pageField must differ from rowField, columnField and the value columns')
  }

  const records: number[][] = []
  for (let row = top + 1; row <= bottom; row += 1) {
    const record: number[] = []
    let valid = true
    for (let col = left; col <= right; col += 1) {
      const field = fields[indexes.get(col)!]!
      const raw = plainContent(cellContentOf(parsed.sheet.getCell(row, col)))
      if (field.numeric) {
        const numeric = parseNumericText(raw)
        if (numeric === null || !Number.isFinite(numeric)) {
          valid = false
          record.push(0)
        } else {
          record.push(numeric)
          field.min = Math.min(field.min ?? 0, numeric)
          field.max = Math.max(field.max ?? 0, numeric)
        }
      } else {
        const items = field.items!
        let index = items.indexOf(raw)
        if (index < 0) {
          index = items.length
          items.push(raw)
        }
        record.push(index)
      }
    }
    if (valid) records.push(record)
  }

  let pageItemIndex: number | undefined
  if (pageFieldIndex !== undefined) {
    const items = fields[pageFieldIndex]!.items ?? []
    pageItemIndex = operation.pageItem === undefined ? 0 : items.indexOf(operation.pageItem)
    if (pageItemIndex < 0) {
      throw new Error(`pivot pageItem "${operation.pageItem}" is not present in ${operation.pageField}`)
    }
  }

  const scopedRecords = pageFieldIndex === undefined
    ? records
    : records.filter((record) => record[pageFieldIndex!] === pageItemIndex)
  const rowItems: number[] = []
  const columnItems: number[] = []
  const groups = new Map<string, number[][]>()
  for (const record of scopedRecords) {
    const rowKey = record[rowFieldIndex]!
    const colKey = columnFieldIndex === undefined ? -1 : record[columnFieldIndex]!
    if (!rowItems.includes(rowKey)) rowItems.push(rowKey)
    if (columnFieldIndex !== undefined && !columnItems.includes(colKey)) columnItems.push(colKey)
    const key = `${rowKey}:${colKey}`
    if (!groups.has(key)) groups.set(key, measureInputs.map(() => []))
    const buckets = groups.get(key)!
    measureInputs.forEach((measure, index) => buckets[index]!.push(record[indexes.get(measure.column)!]!))
  }
  const measures: PivotMeasure[] = measureInputs.map((measure, index) => {
    if (columnFieldIndex === undefined) {
      return {
        name: measure.name,
        fieldIndex: indexes.get(measure.column)!,
        aggregate: measure.aggregate,
        values: rowItems.map((key) => aggregate(groups.get(`${key}:-1`)![index]!, measure.aggregate)),
      }
    }
    return {
      name: measure.name,
      fieldIndex: indexes.get(measure.column)!,
      aggregate: measure.aggregate,
      values: rowItems.map((key) => aggregate(
        columnItems.flatMap((col) => groups.get(`${key}:${col}`)?.[index] ?? []), measure.aggregate,
      )),
      matrix: rowItems.map((key) => columnItems.map((col) =>
        aggregate(groups.get(`${key}:${col}`)?.[index] ?? [], measure.aggregate))),
    }
  })

  const outputName = operation.outputSheet ?? `${parsed.sheet.name}透视`
  let output = findSheet(workbook, outputName)
  if (!output) output = workbook.addWorksheet(outputName)
  const lastOutputCol = 1 + measures.length * (columnFieldIndex === undefined ? 1 : Math.max(columnItems.length, 1))
  for (let row = 1; row <= output.rowCount; row += 1) {
    for (let col = 1; col <= lastOutputCol; col += 1) output.getCell(row, col).value = null
  }
  output.getCell(1, 1).value = rowField.name
  if (columnFieldIndex === undefined) {
    measures.forEach((measure, index) => {
      output.getCell(1, index + 2).value = `${labels[measure.aggregate]}:${measure.name}`
    })
  } else {
    columnItems.forEach((colItem, colIndex) => {
      measures.forEach((measure, measureIndex) => {
        const label = `${columnField!.items![colItem]} - ${labels[measure.aggregate]}:${measure.name}`
        output.getCell(1, 2 + colIndex * measures.length + measureIndex).value = label
      })
    })
  }
  for (let col = 1; col <= lastOutputCol; col += 1) output.getCell(1, col).font = { bold: true }
  rowItems.forEach((key, rowIndex) => {
    output.getCell(rowIndex + 2, 1).value = rowField.items![key]
    measures.forEach((measure, measureIndex) => {
      if (columnFieldIndex === undefined) {
        output.getCell(rowIndex + 2, measureIndex + 2).value = measure.values[rowIndex]!
      } else {
        columnItems.forEach((_colItem, colIndex) => {
          output.getCell(rowIndex + 2, 2 + colIndex * measures.length + measureIndex).value =
            measure.matrix![rowIndex]![colIndex]!
        })
      }
    })
  })
  for (let col = 1; col <= lastOutputCol; col += 1) output.getColumn(col).width = 18

  const sourceRef = `${numberToColumn(left)}${top}:${numberToColumn(right)}${bottom}`
  return {
    name: `PivotTable${cacheId}`,
    cacheId,
    partSuffix,
    sourceSheet: parsed.sheet.name,
    sourceRef,
    outputSheet: outputName,
    outputRef: `${outputName}!A1:${numberToColumn(lastOutputCol)}${rowItems.length + 1}`,
    fields,
    rowFieldIndex,
    rowItemIndices: rowItems,
    columnFieldIndex,
    columnItemIndices: columnFieldIndex === undefined ? undefined : columnItems,
    pageFieldIndex,
    pageItemIndex,
    measures,
    records,
  }
}
