import ExcelJS from 'exceljs'
import { readFile } from 'node:fs/promises'
import { validate, type ValidationResult } from './validator.ts'

export function cellContent(cell: ExcelJS.Cell): string | null {
  if (cell.formula) return `=${cell.formula}`
  const value = cell.value
  if (value === null || value === undefined) return null
  if (typeof value === 'object') {
    if (value instanceof Date) return value.toISOString()
    const text = (value as { text?: unknown }).text
    if (typeof text === 'string') return text
    const richText = (value as { richText?: unknown }).richText
    if (Array.isArray(richText)) return cell.text ?? null
    return JSON.stringify(value)
  }
  return String(value)
}

export async function readWorkbookCells(data: Uint8Array): Promise<Record<string, string>> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(data as any)
  const cells: Record<string, string> = {}
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const content = cellContent(cell)
        if (content !== null) cells[`${sheet.name}!${cell.address}`] = content
      })
    })
  })
  return cells
}

export async function validateWorkbookFile(path: string): Promise<ValidationResult> {
  const data = await readFile(path)
  return validate(await readWorkbookCells(data))
}
