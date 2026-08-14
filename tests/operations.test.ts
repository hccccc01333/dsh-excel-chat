import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  applyOperationsToWorkbook,
  operateWorkbookFile,
  shiftFormulaReferences,
  type ExcelOperation,
} from '../src/operations.ts'
import { readWorkbookCells } from '../src/workbook.ts'

async function makeWorkbook(ops: (workbook: ExcelJS.Workbook) => void): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-operations-'))
  const path = join(dir, 'book.xlsx')
  const workbook = new ExcelJS.Workbook()
  ops(workbook)
  await writeFile(path, await workbook.xlsx.writeBuffer())
  return path
}

async function readAfter(path: string, operations: ExcelOperation[], outName = 'out.xlsx'): Promise<Record<string, string>> {
  const outPath = join(join(path, '..'), outName)
  await applyOperationsToWorkbook(path, operations, outPath)
  return readWorkbookCells(await readFile(outPath))
}

function formulaFixture(): (workbook: ExcelJS.Workbook) => void {
  return (workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('B2').value = 10
    sheet.getCell('C2').value = 4
    sheet.getCell('D2').value = { formula: 'B2-C2' }
    sheet.getCell('B3').value = 20
    sheet.getCell('C3').value = 5
    sheet.getCell('D3').value = { formula: 'B3-C3' }
    sheet.getCell('B4').value = 30
    sheet.getCell('C4').value = 6
    sheet.getCell('D4').value = { formula: 'B4-C4' }
  }
}

test('set writes values and formulas into a workbook copy', async () => {
  const path = await makeWorkbook(formulaFixture())
  const cells = await readAfter(path, [{
    op: 'set',
    cells: { 'Sheet1!D4': '=B4-C3', 'Sheet1!E1': 'total' },
  }])
  assert.equal(cells['Sheet1!D4'], '=B4-C3')
  assert.equal(cells['Sheet1!E1'], 'total')
})

test('fill down copies a formula and shifts relative rows, keeping absolute rows', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('D2').value = { formula: 'B2-$C$2' }
  })
  const cells = await readAfter(path, [{ op: 'fill', source: 'Sheet1!D2', target: 'Sheet1!D2:D4' }])
  assert.equal(cells['Sheet1!D2'], '=B2-$C$2')
  assert.equal(cells['Sheet1!D3'], '=B3-$C$2')
  assert.equal(cells['Sheet1!D4'], '=B4-$C$2')
})

test('fill right shifts relative columns', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('B2').value = { formula: 'A2*2' }
  })
  const cells = await readAfter(path, [{ op: 'fill', source: 'Sheet1!B2', target: 'Sheet1!B2:D2' }])
  assert.equal(cells['Sheet1!B2'], '=A2*2')
  assert.equal(cells['Sheet1!C2'], '=B2*2')
  assert.equal(cells['Sheet1!D2'], '=C2*2')
})

test('insertRows moves cells and shifts same-sheet formula references', async () => {
  const path = await makeWorkbook(formulaFixture())
  const cells = await readAfter(path, [{ op: 'insertRows', sheet: 'Sheet1', row: 3, count: 1 }])
  assert.equal(cells['Sheet1!D2'], '=B2-C2')
  assert.equal(cells['Sheet1!D3'], undefined) // inserted empty row
  assert.equal(cells['Sheet1!D4'], '=B4-C4')  // old D3 moved down, references shifted
  assert.equal(cells['Sheet1!D5'], '=B5-C5')  // old D4 moved down
})

test('deleteRows removes rows and shifts references below the deletion', async () => {
  const path = await makeWorkbook(formulaFixture())
  const cells = await readAfter(path, [{ op: 'deleteRows', sheet: 'Sheet1', row: 3, count: 1 }])
  assert.equal(cells['Sheet1!D2'], '=B2-C2')
  assert.equal(cells['Sheet1!D3'], '=B3-C3') // old D4
  assert.equal(cells['Sheet1!D4'], undefined)
})

test('insertRows also shifts cross-sheet references into the edited sheet', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet1 = workbook.addWorksheet('Sheet1')
    sheet1.getCell('B2').value = 1
    const sheet2 = workbook.addWorksheet('Sheet2')
    sheet2.getCell('E2').value = { formula: 'Sheet1!B2*2' }
    sheet2.getCell('E3').value = { formula: 'Sheet1!B3*2' }
  })
  const cells = await readAfter(path, [{ op: 'insertRows', sheet: 'Sheet1', row: 2, count: 1 }])
  assert.equal(cells['Sheet2!E2'], '=Sheet1!B3*2')
  assert.equal(cells['Sheet2!E3'], '=Sheet1!B4*2')
})

test('renameSheet updates references across the workbook', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet1 = workbook.addWorksheet('Sheet1')
    sheet1.getCell('B2').value = 1
    const sheet2 = workbook.addWorksheet('Sheet2')
    sheet2.getCell('E2').value = { formula: 'Sheet1!B2*2' }
  })
  const cells = await readAfter(path, [{ op: 'renameSheet', oldName: 'Sheet1', newName: 'Data' }])
  assert.equal(cells['Sheet2!E2'], '=Data!B2*2')
  assert.ok(!Object.keys(cells).some((id) => id.startsWith('Sheet1!')))
  assert.ok(Object.keys(cells).some((id) => id.startsWith('Data!')))
})

test('addSheet and deleteSheet change the workbook sheet set', async () => {
  const path = await makeWorkbook(formulaFixture())
  const addedPath = join(join(path, '..'), 'added.xlsx')
  await applyOperationsToWorkbook(path, [{ op: 'addSheet', name: 'Notes' }], addedPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(addedPath)
  assert.ok(workbook.getWorksheet('Notes'))

  const removedPath = join(join(path, '..'), 'removed.xlsx')
  await applyOperationsToWorkbook(path, [{ op: 'deleteSheet', name: 'Sheet1' }], removedPath)
  const after = new ExcelJS.Workbook()
  await after.xlsx.readFile(removedPath)
  assert.equal(after.getWorksheet('Sheet1'), undefined)
})

test('merge and unmerge toggle the merged range', async () => {
  const path = await makeWorkbook(formulaFixture())
  const mergedPath = join(join(path, '..'), 'merged.xlsx')
  await applyOperationsToWorkbook(path, [{ op: 'merge', range: 'Sheet1!A1:B2' }], mergedPath)
  const merged = new ExcelJS.Workbook()
  await merged.xlsx.readFile(mergedPath)
  assert.deepEqual(merged.getWorksheet('Sheet1')!.model.merges, ['A1:B2'])

  const unmergedPath = join(join(path, '..'), 'unmerged.xlsx')
  await applyOperationsToWorkbook(path, [{ op: 'unmerge', range: 'Sheet1!A1:B2' }], unmergedPath)
  const unmerged = new ExcelJS.Workbook()
  await unmerged.xlsx.readFile(unmergedPath)
  assert.deepEqual(unmerged.getWorksheet('Sheet1')!.model.merges, [])
})

test('clear empties the selected cells', async () => {
  const path = await makeWorkbook(formulaFixture())
  const cells = await readAfter(path, [{ op: 'clear', cells: ['Sheet1!D3', 'Sheet1!B3'] }])
  assert.equal(cells['Sheet1!D3'], undefined)
  assert.equal(cells['Sheet1!B3'], undefined)
  assert.equal(cells['Sheet1!D2'], '=B2-C2')
})

test('operateWorkbookFile returns post-operation validation and flags broken patterns', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'edited.xlsx')
  const result = await operateWorkbookFile(path, [
    { op: 'set', cells: { 'Sheet1!D4': '=B4-C3' } },
  ], outPath)
  assert.equal(result.outputPath, outPath)
  assert.ok(result.validation.anomalies.some((anomaly) => anomaly.kind === 'reference-offset' && anomaly.cell === 'Sheet1!D4'))
})

test('shiftFormulaReferences moves relative rows only when gated by the threshold', () => {
  assert.equal(shiftFormulaReferences('=B2-C2', 'Sheet1', 'Sheet1', { rowDelta: 1, rowThreshold: 3 }), '=B2-C2')
  assert.equal(shiftFormulaReferences('=B3-C3', 'Sheet1', 'Sheet1', { rowDelta: 1, rowThreshold: 3 }), '=B4-C4')
  // Cross-sheet reference stays; same-sheet reference shifts.
  assert.equal(shiftFormulaReferences('=Sheet2!B3-C3', 'Sheet1', 'Sheet1', { rowDelta: 1, rowThreshold: 3 }), '=Sheet2!B3-C4')
})
