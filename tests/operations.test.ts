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

test('set types numbers, dates, and booleans instead of writing text', async () => {
  const path = await makeWorkbook((workbook) => {
    workbook.addWorksheet('Sheet1')
  })
  const outPath = join(join(path, '..'), 'typed.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'set',
    cells: {
      'Sheet1!A1': '100',
      'Sheet1!A2': '2026-01-15',
      'Sheet1!A3': 'true',
      'Sheet1!A4': 'abc',
    },
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('A1').value, 100)
  assert.ok(sheet.getCell('A2').value instanceof Date)
  assert.equal(sheet.getCell('A3').value, true)
  assert.equal(sheet.getCell('A4').value, 'abc')
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

test('insertColumns shifts formula references to the right', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A2').value = 1
    sheet.getCell('B2').value = { formula: 'A2*2' }
    sheet.getCell('C2').value = { formula: 'B2*2' }
  })
  const cells = await readAfter(path, [{ op: 'insertColumns', sheet: 'Sheet1', column: 'B', count: 1 }])
  assert.equal(cells['Sheet1!B2'], undefined) // inserted empty column
  assert.equal(cells['Sheet1!C2'], '=A2*2')   // old B2
  assert.equal(cells['Sheet1!D2'], '=C2*2')   // old C2, reference shifted
})

test('deleteColumns removes columns and shifts references after the deletion', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A2').value = 1
    sheet.getCell('B2').value = { formula: 'A2*2' }
    sheet.getCell('C2').value = { formula: 'B2*2' }
    sheet.getCell('D2').value = { formula: 'C2*2' }
  })
  const outPath = join(join(path, '..'), 'delcol.xlsx')
  const result = await applyOperationsToWorkbook(path, [
    { op: 'deleteColumns', sheet: 'Sheet1', column: 'B', count: 1 },
  ], outPath)
  assert.ok(result.warnings.some((warning) => /references a deleted column/.test(warning.message)))
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('B2').formula, '#REF!*2') // old C2, reference into deleted B became #REF!
  assert.equal(sheet.getCell('C2').formula, 'B2*2')  // old D2, reference shifted C -> B
  assert.equal(sheet.getCell('D2').value, null)
})

test('deleteRows turns references into deleted rows into #REF!', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('D2').value = { formula: 'B2-C2' }
    sheet.getCell('D3').value = { formula: 'B3-C3' }
    sheet.getCell('E2').value = { formula: 'D2+D3' }
  })
  const outPath = join(join(path, '..'), 'ref.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'deleteRows', sheet: 'Sheet1', row: 3, count: 1 },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('E2').formula, 'D2+#REF!')
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

test('copyRange copies values and shifts formulas by the destination offset', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'copied.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'copyRange', source: 'Sheet1!B2:D3', target: 'Sheet1!E2' },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('E2').value, 10)
  assert.equal(sheet.getCell('F2').value, 4)
  assert.equal(sheet.getCell('G2').formula, 'E2-F2') // formula shifted +3 columns
  assert.equal(sheet.getCell('G3').formula, 'E3-F3')
})

test('copyRange with move clears the source range', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'moved.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'copyRange', source: 'Sheet1!B2:C2', target: 'Sheet1!E2', move: true },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('B2').value, null)
  assert.equal(sheet.getCell('E2').value, 10)
})

test('fillSeries fills numeric and date sequences', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 1
    sheet.getCell('C1').value = new Date(2026, 0, 1)
  })
  const outPath = join(join(path, '..'), 'series.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'fillSeries', start: 'Sheet1!A1', target: 'Sheet1!A1:A5' },
    { op: 'fillSeries', start: 'Sheet1!C1', target: 'Sheet1!C1:C3' },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('A5').value, 5)
  const third = sheet.getCell('C3').value as Date
  assert.equal(third.getFullYear(), 2026)
  assert.equal(third.getMonth(), 0)
  assert.equal(third.getDate(), 3)
})

test('style applies font, fill, number format, and alignment to a range', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'styled.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'style',
    range: 'Sheet1!A1:B2',
    style: {
      bold: true,
      fontColor: 'FF0000',
      fill: 'FFFF00',
      numberFormat: '#,##0.00',
      hAlign: 'center',
      wrapText: true,
    },
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const cell = workbook.getWorksheet('Sheet1')!.getCell('A1')
  assert.equal(cell.font.bold, true)
  assert.equal(cell.font.color.argb, 'FFFF0000')
  assert.equal((cell.fill as { fgColor?: { argb?: string } }).fgColor?.argb, 'FFFFFF00')
  assert.equal(cell.numFmt, '#,##0.00')
  assert.equal(cell.alignment.horizontal, 'center')
  assert.equal(cell.alignment.wrapText, true)
})

test('setColumnWidth, setRowHeight, and freezePanes update sheet layout', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'layout.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'setColumnWidth', sheet: 'Sheet1', column: 'B', width: 24 },
    { op: 'setRowHeight', sheet: 'Sheet1', row: 1, height: 30 },
    { op: 'freezePanes', sheet: 'Sheet1', row: 2, column: 'B' },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getColumn('B').width, 24)
  assert.equal(sheet.getRow(1).height, 30)
  assert.equal(sheet.views[0]!.state, 'frozen')
  assert.equal(sheet.views[0]!.ySplit, 1)
})

test('findReplace swaps text across cells and reports the count', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 'old name'
    sheet.getCell('A2').value = 'OLD NAME'
    sheet.getCell('A3').value = 'keep'
  })
  const outPath = join(join(path, '..'), 'replaced.xlsx')
  const result = await applyOperationsToWorkbook(path, [
    { op: 'findReplace', find: 'old', replace: 'new', sheet: 'Sheet1', matchCase: false },
  ], outPath)
  assert.equal(result.warnings[0]!.message, 'findReplace replaced 2 occurrence(s)')
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('A1').value, 'new name')
  assert.equal(sheet.getCell('A2').value, 'new NAME')
  assert.equal(sheet.getCell('A3').value, 'keep')
})

test('duplicateSheet copies values and merges; hideSheet and setTabColor update the sheet', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 1
    sheet.mergeCells('A1:B1')
  })
  const outPath = join(join(path, '..'), 'sheets.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'duplicateSheet', name: 'Sheet1', newName: 'Copy' },
    { op: 'hideSheet', name: 'Copy' },
    { op: 'setTabColor', name: 'Copy', color: '00AA00' },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const copy = workbook.getWorksheet('Copy')!
  assert.equal(copy.getCell('A1').value, 1)
  assert.deepEqual(copy.model.merges, ['A1:B1'])
  assert.equal(copy.state, 'hidden')
  assert.equal(copy.properties.tabColor?.argb, 'FF00AA00')
})

test('sortRange sorts by one key descending and keeps the header row in place', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 'Name'
    sheet.getCell('B1').value = 'Score'
    sheet.getCell('A2').value = 'Alice'
    sheet.getCell('B2').value = 90
    sheet.getCell('A3').value = 'Bob'
    sheet.getCell('B3').value = 70
    sheet.getCell('A4').value = 'Carol'
    sheet.getCell('B4').value = 80
  })
  const outPath = join(join(path, '..'), 'sorted.xlsx')
  const result = await applyOperationsToWorkbook(path, [
    { op: 'sortRange', range: 'Sheet1!A1:B4', keys: [{ column: 'B', direction: 'desc' }], headerRows: 1 },
  ], outPath)
  assert.ok(result.warnings.some((warning) => warning.message.startsWith('sortRange')))
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('A1').value, 'Name')
  assert.equal(sheet.getCell('A2').value, 'Alice')
  assert.equal(sheet.getCell('A3').value, 'Carol')
  assert.equal(sheet.getCell('A4').value, 'Bob')
})

test('dataValidation adds a dropdown list and numeric between rules', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'validated.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'dataValidation', range: 'Sheet1!A2:A4', type: 'list', formula1: '高,中,低', allowBlank: true },
    {
      op: 'dataValidation',
      range: 'Sheet1!B2:B4',
      type: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '100',
      error: '必须填写 1-100 的整数',
      errorTitle: '输入错误',
    },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('A2').dataValidation?.type, 'list')
  assert.deepEqual(sheet.getCell('A2').dataValidation?.formulae, ['"高,中,低"'])
  assert.equal(sheet.getCell('B3').dataValidation?.type, 'whole')
  assert.equal(sheet.getCell('B3').dataValidation?.operator, 'between')
  assert.deepEqual(sheet.getCell('B3').dataValidation?.formulae, [1, 100])
  assert.equal(sheet.getCell('B3').dataValidation?.error, '必须填写 1-100 的整数')
})

test('conditionalFormatting adds a cellIs rule with a red fill', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('B2').value = 90
    sheet.getCell('B3').value = 70
    sheet.getCell('B4').value = 80
  })
  const outPath = join(join(path, '..'), 'cf.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'conditionalFormatting',
    range: 'Sheet1!B2:B4',
    rules: [{
      type: 'cellIs',
      operator: 'greaterThan',
      formula: 80,
      style: { fill: 'FF0000', bold: true, fontColor: 'FFFFFF' },
    }],
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  const cf = (sheet as unknown as { conditionalFormattings?: Array<{ ref: string; rules: Array<{ type: string; operator: string; style?: { fill?: { bgColor?: { argb?: string } } } }> }> }).conditionalFormattings
  const entry = cf?.[0]
  assert.ok(entry, 'conditional formatting missing')
  assert.equal(entry!.ref, 'B2:B4')
  assert.equal(entry!.rules[0]!.type, 'cellIs')
  assert.equal(entry!.rules[0]!.operator, 'greaterThan')
  assert.equal(entry!.rules[0]!.style?.fill?.bgColor?.argb, 'FFFF0000')
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
