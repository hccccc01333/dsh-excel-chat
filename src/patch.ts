import ExcelJS from 'exceljs'
import { readFile } from 'node:fs/promises'
import { cellContent, stripPivotTableParts } from './workbook.ts'

export interface CellPatch {
  id: string
  kind: 'formula' | 'value'
  oldValue: string
  newValue: string
}

export function applyPatches(cells: Record<string, string>, patches: CellPatch[]): Record<string, string> {
  const result: Record<string, string> = { ...cells }
  for (const patch of patches) {
    const current = result[patch.id] ?? ''
    if (current !== patch.oldValue) {
      throw new Error(`patch precondition failed for ${patch.id}: expected ${patch.oldValue}, got ${current}`)
    }
    if (patch.newValue === '') delete result[patch.id]
    else result[patch.id] = patch.newValue
  }
  return result
}

export function revertPatches(cells: Record<string, string>, patches: CellPatch[]): Record<string, string> {
  return applyPatches(
    cells,
    patches.map((patch) => ({ ...patch, oldValue: patch.newValue, newValue: patch.oldValue })),
  )
}

export async function applyPatchesToWorkbook(
  inputPath: string,
  patches: CellPatch[],
  outputPath: string = inputPath,
): Promise<void> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(stripPivotTableParts(await readFile(inputPath)) as any)
  for (const patch of patches) {
    const bang = patch.id.lastIndexOf('!')
    const sheetName = bang >= 0
      ? patch.id.slice(0, bang).replace(/^'|'$/g, '')
      : workbook.worksheets[0]?.name
    const cellPart = bang >= 0 ? patch.id.slice(bang + 1) : patch.id
    const cellMatch = /^([A-Za-z]{1,3})(\d+)$/.exec(cellPart)
    if (!sheetName || !cellMatch) {
      throw new Error(`patch failed for ${patch.id}: invalid cell id`)
    }
    const sheet = workbook.getWorksheet(sheetName)
    if (!sheet) throw new Error(`patch failed for ${patch.id}: sheet ${sheetName} not found`)
    const target = sheet.getCell(`${cellMatch[1]}${cellMatch[2]}`)
    const current = cellContent(target)
    if ((current ?? '') !== patch.oldValue) {
      throw new Error(`patch precondition failed for ${patch.id}: expected ${patch.oldValue}, got ${current}`)
    }
    target.value = patch.newValue === ''
      ? null
      : patch.newValue.startsWith('=')
        ? { formula: patch.newValue.slice(1) }
        : patch.newValue
  }
  await workbook.xlsx.writeFile(outputPath)
}
