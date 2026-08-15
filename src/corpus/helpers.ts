import ExcelJS from 'exceljs'
import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'

export interface CorpusSheet {
  name: string
  headers: string[]
  rows: Array<Array<string | number | boolean | null>>
}

/** Build a small realistic workbook from plain sheet descriptors. */
export async function buildCorpusWorkbook(
  dir: string,
  id: string,
  sheets: CorpusSheet[],
): Promise<string> {
  const workbook = new ExcelJS.Workbook()
  for (const spec of sheets) {
    const ws = workbook.addWorksheet(spec.name)
    spec.headers.forEach((header, index) => {
      ws.getCell(`${columnLetter(index)}1`).value = header
    })
    spec.rows.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        if (value === null || value === undefined) return
        const cell = ws.getCell(`${columnLetter(colIndex)}${rowIndex + 2}`)
        cell.value = typeof value === 'string' && value.startsWith('=') ? { formula: value.slice(1) } : value
      })
    })
  }
  const path = join(dir, `${id}.xlsx`)
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

export function columnLetter(index: number): string {
  return String.fromCharCode(65 + index)
}
