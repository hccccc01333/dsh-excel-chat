import ExcelJS from 'exceljs'
import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import type { ExcelStyle } from '../operation-types.ts'
import { excelStyleToWorkbookStyle } from '../op/formatting.ts'

export interface CorpusSheet {
  name: string
  headers: string[]
  rows: Array<Array<string | number | boolean | null>>
  /**
   * Pre-existing formatting, keyed by cell id (`B2`) or by `*` for the whole sheet.
   *
   * The builder could only set values, so no task could start from a styled cell —
   * which is why the corpus asserted formatting on 5 of 185 checks and why operations
   * like `copyStyle` had nowhere to be exercised. Styles go through the same
   * conversion the `style` operation uses, so the corpus and the plugin cannot
   * disagree about what `hAlign: 'center'` means.
   */
  styles?: Record<string, ExcelStyle>
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
    for (const [target, style] of Object.entries(spec.styles ?? {})) {
      const converted = excelStyleToWorkbookStyle(style)
      if (target === '*') {
        ws.eachRow((row) => row.eachCell((cell) => { cell.style = converted }))
        continue
      }
      ws.getCell(target).style = converted
    }
  }
  const path = join(dir, `${id}.xlsx`)
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

export function columnLetter(index: number): string {
  return String.fromCharCode(65 + index)
}
