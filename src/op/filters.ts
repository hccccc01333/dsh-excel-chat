/**
 * Selecting rows: filtering to a range, testing a criterion, and mail merge.
 *
 * The criterion matcher is shared by the filter and the merge so the two cannot
 * disagree about what `contains` means. Split out of `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs'
import { t } from '../i18n.ts'
import { columnToNumber, numberToColumn } from '../formula.ts'
import type { ExcelOperation } from '../operation-types.ts'
import { cellContent } from '../workbook.ts'
import { findSheet, parseTargetCell, parseRange } from './core.ts'
export function applyFilterToRange(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'filterToRange' }>,
): void {
  const parsed = parseRange(workbook, options.source)
  const target = parseTargetCell(workbook, options.target, parsed.sheet.name)
  const matchAll = options.matchAll ?? true

  const headerRow: Array<ExcelJS.CellValue> = []
  for (let col = parsed.startCol; col <= parsed.endCol; col++) {
    headerRow.push(parsed.sheet.getCell(`${numberToColumn(col)}${parsed.startRow}`).value)
  }
  let targetRow = target.row
  headerRow.forEach((value, index) => {
    target.sheet.getCell(`${numberToColumn(target.col + index)}${targetRow}`).value = value
  })
  targetRow += 1

  for (let row = parsed.startRow + 1; row <= parsed.endRow; row++) {
    let matched = matchAll
    for (const criterion of options.criteria) {
      const col = columnToNumber(criterion.column)
      const actual = parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value
      const ok = matchesCriterion(actual, criterion.operator, criterion.value)
      if (matchAll && !ok) {
        matched = false
        break
      }
      if (!matchAll && ok) {
        matched = true
        break
      }
    }
    if (!matched) continue
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
      target.sheet.getCell(`${numberToColumn(target.col + (col - parsed.startCol))}${targetRow}`).value =
        parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value
    }
    targetRow += 1
  }
}

export function matchesCriterion(
  actual: ExcelJS.CellValue,
  operator: 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains',
  expected: string | number,
): boolean {
  const actualNumber = typeof actual === 'number' ? actual : null
  const expectedNumber = typeof expected === 'number' ? expected : Number(expected)
  const actualText = actual === null || actual === undefined ? '' : String(actual)
  const expectedText = String(expected)
  switch (operator) {
    case 'eq':
      return actualNumber !== null && Number.isFinite(expectedNumber)
        ? actualNumber === expectedNumber
        : actualText.toLowerCase() === expectedText.toLowerCase()
    case 'neq':
      return !matchesCriterion(actual, 'eq', expected)
    case 'contains':
      return actualText.toLowerCase().includes(expectedText.toLowerCase())
    case 'gt':
      return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber > expectedNumber
    case 'gte':
      return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber >= expectedNumber
    case 'lt':
      return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber < expectedNumber
    case 'lte':
      return actualNumber !== null && Number.isFinite(expectedNumber) && actualNumber <= expectedNumber
  }
}

export function applyMailMerge(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'mailMerge' }>,
): void {
  const template = parseRange(workbook, options.template)
  const data = parseRange(workbook, options.data)
  const headers = new Map<string, number>()
  for (let col = data.startCol; col <= data.endCol; col++) {
    const raw = data.sheet.getCell(`${numberToColumn(col)}${data.startRow}`).value
    headers.set(String(raw ?? '').toLowerCase(), col)
  }
  const templateRows: Array<Record<number, ExcelJS.CellValue>> = []
  for (let row = template.startRow; row <= template.endRow; row++) {
    const cells: Record<number, ExcelJS.CellValue> = {}
    for (let col = template.startCol; col <= template.endCol; col++) {
      cells[col] = template.sheet.getCell(`${numberToColumn(col)}${row}`).value
    }
    templateRows.push(cells)
  }

  const outputSheetName = options.outputSheet ?? `${template.sheet.name}-合并`
  let output = findSheet(workbook, outputSheetName)
  if (!output) output = workbook.addWorksheet(outputSheetName)
  let outputRow = 1
  const placeholder = /\{([^{}]+)\}/g
  for (let dataRow = data.startRow + 1; dataRow <= data.endRow; dataRow++) {
    const record = new Map<string, ExcelJS.CellValue>()
    for (const [header, col] of headers) {
      record.set(header, data.sheet.getCell(`${numberToColumn(col)}${dataRow}`).value)
    }
    for (const templateRow of templateRows) {
      for (const [col, value] of Object.entries(templateRow)) {
        const column = Number(col)
        const text = value === null || value === undefined ? '' : String(value)
        if (/^\{[^{}]+\}$/.test(text.trim())) {
          const key = text.trim().slice(1, -1).toLowerCase()
          output.getCell(`${numberToColumn(column)}${outputRow}`).value = record.get(key) ?? text
          continue
        }
        if (placeholder.test(text)) {
          placeholder.lastIndex = 0
          output.getCell(`${numberToColumn(column)}${outputRow}`).value = text.replace(placeholder, (_match, name: string) => {
            const replacement = record.get(String(name).toLowerCase())
            return replacement === undefined ? _match : String(replacement)
          })
          continue
        }
        output.getCell(`${numberToColumn(column)}${outputRow}`).value = value
      }
      outputRow += 1
    }
  }
}
