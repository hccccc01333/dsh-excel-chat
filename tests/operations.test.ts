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
import { readPatchLog, rollbackPatchLog } from '../src/diff.ts'
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

test('style supports font size/name and per-side borders', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'bordered.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'style',
    range: 'Sheet1!A1:B2',
    style: {
      fontSize: 14,
      fontName: '微软雅黑',
      border: {
        top: { style: 'thin', color: 'FF0000' },
        bottom: { style: 'double', color: '0000FF' },
      },
    },
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const cell = workbook.getWorksheet('Sheet1')!.getCell('A1')
  assert.equal(cell.font.size, 14)
  assert.equal(cell.font.name, '微软雅黑')
  assert.equal(cell.border?.top?.style, 'thin')
  assert.equal((cell.border?.top?.color as { argb?: string }).argb, 'FFFF0000')
  assert.equal(cell.border?.bottom?.style, 'double')
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

test('conditionalFormatting supports dataBar, containsText, and top10 rules', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('B2').value = 90
    sheet.getCell('B3').value = 70
    sheet.getCell('B4').value = 80
    sheet.getCell('A2').value = '正常'
    sheet.getCell('A3').value = '异常'
    sheet.getCell('A4').value = '正常'
  })
  const outPath = join(join(path, '..'), 'cf-rich.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'conditionalFormatting',
    range: 'Sheet1!B2:B4',
    rules: [
      { type: 'dataBar', color: '63BE7B' },
      { type: 'top10', rank: 2 },
    ],
  }, {
    op: 'conditionalFormatting',
    range: 'Sheet1!A2:A4',
    rules: [{ type: 'containsText', text: '异常', style: { fill: 'FF0000' } }],
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  const cf = (sheet as unknown as { conditionalFormattings: Array<{ rules: Array<{ type: string }> }> }).conditionalFormattings
  const ruleTypes = cf.flatMap((entry) => entry.rules.map((rule) => rule.type))
  assert.ok(ruleTypes.includes('dataBar'))
  assert.ok(ruleTypes.includes('top10'))
  assert.ok(ruleTypes.includes('containsText'))
})

test('conditionalFormatting supports duplicate, blank, error, average, and timePeriod rules', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A2').value = 'x'
    sheet.getCell('A3').value = 'x'
    sheet.getCell('B2').value = 1
    sheet.getCell('B3').value = 2
  })
  const outPath = join(join(path, '..'), 'cf-more.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'conditionalFormatting',
    range: 'Sheet1!A2:A4',
    rules: [
      { type: 'duplicateValues', style: { fill: 'FFEB84' } },
      { type: 'blanks', style: { fill: 'F8696B' } },
      { type: 'errors', style: { fill: 'FF0000' } },
    ],
  }, {
    op: 'conditionalFormatting',
    range: 'Sheet1!B2:B4',
    rules: [
      { type: 'aboveAverage' },
      { type: 'timePeriod', timePeriod: 'today' },
      { type: 'notContainsText', text: 'x' },
    ],
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  const cf = (sheet as unknown as { conditionalFormattings: Array<{ rules: Array<{ type: string; formulae?: string[] }> }> }).conditionalFormattings
  const types = cf.flatMap((entry) => entry.rules.map((rule) => rule.type))
  assert.ok(types.includes('aboveAverage'))
  assert.ok(types.includes('timePeriod'))
  const formulae = cf.flatMap((entry) => entry.rules.flatMap((rule) => rule.formulae ?? []))
  assert.ok(formulae.some((formula) => formula.includes('ISBLANK')))
  assert.ok(formulae.some((formula) => formula.includes('ISERROR')))
  assert.ok(formulae.some((formula) => formula.includes('COUNTIF') && formula.includes('>1')))
  assert.ok(formulae.some((formula) => formula.includes('SEARCH("x"')))
})

test('autoFilter adds filter dropdowns to the header range', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 'Name'
    sheet.getCell('B1').value = 'Score'
    sheet.getCell('A2').value = 'Alice'
    sheet.getCell('B2').value = 90
  })
  const outPath = join(join(path, '..'), 'filtered.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'autoFilter', range: 'Sheet1!A1:B2' },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const filter = workbook.getWorksheet('Sheet1')!.autoFilter
  assert.equal(filter, 'A1:B2')
})

test('addTable registers a structured table with striped rows', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 'Name'
    sheet.getCell('B1').value = 'Score'
    sheet.getCell('A2').value = 'Alice'
    sheet.getCell('B2').value = 90
    sheet.getCell('A3').value = 'Bob'
    sheet.getCell('B3').value = 70
  })
  const outPath = join(join(path, '..'), 'table.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'addTable',
    name: 'SalesTable',
    range: 'Sheet1!A1:B3',
    showRowStripes: true,
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const tables = Object.values(workbook.getWorksheet('Sheet1')!.tables) as Array<{
    name: string
    model: { tableRef: string }
  }>
  assert.equal(tables.length, 1)
  assert.equal(tables[0]!.name, 'SalesTable')
  assert.equal(tables[0]!.model.tableRef, 'A1:B3')
})

test('subtotal inserts per-group summary rows and a grand total', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = '类别'
    sheet.getCell('B1').value = '金额'
    sheet.getCell('A2').value = 'A'
    sheet.getCell('B2').value = 10
    sheet.getCell('A3').value = 'A'
    sheet.getCell('B3').value = 20
    sheet.getCell('A4').value = 'B'
    sheet.getCell('B4').value = 5
    sheet.getCell('A5').value = 'B'
    sheet.getCell('B5').value = 15
  })
  const outPath = join(join(path, '..'), 'subtotal.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'subtotal',
    sheet: 'Sheet1',
    range: 'Sheet1!A1:B5',
    groupColumn: 'A',
    summaryColumns: [{ column: 'B', function: 'sum' }],
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  assert.equal(sheet.getCell('A4').value, 'A 汇总')
  assert.equal(sheet.getCell('B4').formula, 'SUBTOTAL(9,B2:B3)')
  assert.equal(sheet.getCell('A7').value, 'B 汇总')
  assert.equal(sheet.getCell('B7').formula, 'SUBTOTAL(9,B5:B6)')
  assert.equal(sheet.getCell('A8').value, '总计')
  assert.equal(sheet.getCell('B8').formula, 'SUBTOTAL(9,B2:B7)')
  assert.equal(sheet.getCell('B4').font.bold, true)
})

test('subtotal shifts formulas that reference the edited range', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = '类别'
    sheet.getCell('B1').value = '金额'
    sheet.getCell('A2').value = 'A'
    sheet.getCell('B2').value = 10
    sheet.getCell('A3').value = 'A'
    sheet.getCell('B3').value = 20
    sheet.getCell('A4').value = 'B'
    sheet.getCell('B4').value = 5
    sheet.getCell('A5').value = 'B'
    sheet.getCell('B5').value = 15
    const summary = workbook.addWorksheet('汇总')
    summary.getCell('A1').value = '类别'
    summary.getCell('A2').value = 'A'
    summary.getCell('B2').value = { formula: 'SUMIFS(Sheet1!B2:B5,Sheet1!A2:A5,A2)' }
  })
  const outPath = join(join(path, '..'), 'subtotal-shift.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'subtotal',
    sheet: 'Sheet1',
    range: 'Sheet1!A1:B5',
    groupColumn: 'A',
    summaryColumns: [{ column: 'B', function: 'sum' }],
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  assert.equal(
    workbook.getWorksheet('汇总')!.getCell('B2').formula,
    'SUMIFS(Sheet1!B2:B6,Sheet1!A2:A6,A2)',
  )
})

test('aggregateReport builds a dynamic summary sheet with SUMIFS formulas', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = '类别'
    sheet.getCell('B1').value = '金额'
    sheet.getCell('C1').value = '数量'
    sheet.getCell('A2').value = 'A'
    sheet.getCell('B2').value = 10
    sheet.getCell('C2').value = 2
    sheet.getCell('A3').value = 'A'
    sheet.getCell('B3').value = 20
    sheet.getCell('C3').value = 4
    sheet.getCell('A4').value = 'B'
    sheet.getCell('B4').value = 5
    sheet.getCell('C4').value = 1
  })
  const outPath = join(join(path, '..'), 'report-summary.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'aggregateReport',
    source: 'Sheet1!A1:C4',
    groupColumn: 'A',
    metrics: [
      { column: 'B', function: 'sum' },
      { column: 'C', function: 'average' },
    ],
  }], outPath)
  const cells = await readWorkbookCells(await readFile(outPath))
  assert.equal(cells['Sheet1-汇总!A1'], '类别')
  assert.equal(cells['Sheet1-汇总!B1'], '金额 合计')
  assert.equal(cells['Sheet1-汇总!A2'], 'A')
  assert.equal(cells['Sheet1-汇总!B2'], '=SUMIFS(Sheet1!$B$2:$B$4,Sheet1!$A$2:$A$4,A2)')
  assert.equal(cells['Sheet1-汇总!C2'], '=AVERAGEIFS(Sheet1!$C$2:$C$4,Sheet1!$A$2:$A$4,A2)')
  assert.equal(cells['Sheet1-汇总!A4'], '总计')
  assert.equal(cells['Sheet1-汇总!B4'], '=SUM(B2:B3)')
})

test('filterToRange writes rows matching criteria to a target range', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = '名称'
    sheet.getCell('B1').value = '金额'
    sheet.getCell('A2').value = '苹果'
    sheet.getCell('B2').value = 10
    sheet.getCell('A3').value = '香蕉'
    sheet.getCell('B3').value = 5
    sheet.getCell('A4').value = '苹果'
    sheet.getCell('B4').value = 20
    sheet.getCell('A5').value = '橙子'
    sheet.getCell('B5').value = 8
  })
  const cells = await readAfter(path, [{
    op: 'filterToRange',
    source: 'Sheet1!A1:B5',
    criteria: [{ column: 'A', operator: 'eq', value: '苹果' }],
    target: 'Sheet1!D1',
  }])
  assert.equal(cells['Sheet1!D1'], '名称')
  assert.equal(cells['Sheet1!D2'], '苹果')
  assert.equal(cells['Sheet1!E2'], '10')
  assert.equal(cells['Sheet1!D3'], '苹果')
  assert.equal(cells['Sheet1!E3'], '20')
  assert.equal(cells['Sheet1!D4'], undefined)
})

test('protectSheet and unprotectSheet round-trip through a file', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'protected.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'protectSheet', sheet: 'Sheet1', password: 'secret' },
  ], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const sheet = workbook.getWorksheet('Sheet1')!
  sheet.unprotect('secret')
  sheet.getCell('A1').value = 1 // still writable after unprotect
  assert.equal(sheet.getCell('A1').value, 1)
})

test('pageSetup sets print area, orientation, and margins', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'page.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'pageSetup',
    sheet: 'Sheet1',
    printArea: 'A1:D10',
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    margins: { top: 0.5, left: 0.3 },
    centerHorizontally: true,
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const pageSetup = workbook.getWorksheet('Sheet1')!.pageSetup
  assert.equal(pageSetup.printArea, 'A1:D10')
  assert.equal(pageSetup.orientation, 'landscape')
  assert.equal(pageSetup.fitToPage, true)
  assert.equal(pageSetup.fitToWidth, 1)
  assert.equal(pageSetup.margins?.top, 0.5)
  assert.equal(pageSetup.horizontalCentered, true)
})

test('definedName registers a named range on the workbook', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'named.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'definedName',
    name: 'SalesRange',
    ref: "Sheet1!$A$1:$D$10",
  }], outPath)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(outPath)
  const names = workbook.definedNames.model as Array<{ name: string; ranges: string[] }>
  assert.ok(names.some((entry) => entry.name === 'SalesRange' && entry.ranges.some((range) => range.includes('$A$1:$D$10'))))
})

test('importCsv and exportCsv round-trip with typed values and formula-injection guard', async () => {
  const path = await makeWorkbook((workbook) => {
    const sheet = workbook.addWorksheet('Sheet1')
    sheet.getCell('A1').value = '名称'
    sheet.getCell('B1').value = '金额'
    sheet.getCell('A2').value = '苹果'
    sheet.getCell('B2').value = 100
    sheet.getCell('A3').value = '=注入'
  })
  const dir = join(join(path, '..'))
  const csvPath = join(dir, 'out.csv')
  await applyOperationsToWorkbook(path, [
    { op: 'exportCsv', file: csvPath, sheet: 'Sheet1' },
  ], join(dir, 'exported.xlsx'))
  const csv = await readFile(csvPath, 'utf8')
  assert.ok(csv.includes('名称,金额'))
  assert.ok(csv.includes('苹果,100'))
  assert.ok(csv.includes("'=注入"))

  const importOut = join(dir, 'imported.xlsx')
  await applyOperationsToWorkbook(path, [
    { op: 'importCsv', file: csvPath, sheet: '导入' },
  ], importOut)
  const cells = await readWorkbookCells(await readFile(importOut))
  assert.equal(cells['导入!A1'], '名称')
  assert.equal(cells['导入!B2'], '100')
})

test('mailMerge expands a template with placeholders per data record', async () => {
  const path = await makeWorkbook((workbook) => {
    const template = workbook.addWorksheet('Sheet1')
    template.getCell('A1').value = '{姓名}'
    template.getCell('B1').value = '你好，{金额} 元'
    const data = workbook.addWorksheet('Sheet2')
    data.getCell('A1').value = '姓名'
    data.getCell('B1').value = '金额'
    data.getCell('A2').value = '张三'
    data.getCell('B2').value = 100
    data.getCell('A3').value = '李四'
    data.getCell('B3').value = 200
  })
  const cells = await readAfter(path, [{
    op: 'mailMerge',
    template: 'Sheet1!A1:B1',
    data: 'Sheet2!A1:B3',
    outputSheet: '通知',
  }])
  assert.equal(cells['通知!A1'], '张三')
  assert.equal(cells['通知!B1'], '你好，100 元')
  assert.equal(cells['通知!A2'], '李四')
  assert.equal(cells['通知!B2'], '你好，200 元')
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

test('operateWorkbookFile writes a patch log and rollback restores the original cells', async () => {
  const path = await makeWorkbook(formulaFixture())
  const outPath = join(join(path, '..'), 'undo-edit.xlsx')
  const result = await operateWorkbookFile(path, [
    { op: 'set', cells: { 'Sheet1!D4': '=B4-C3' } },
    { op: 'style', range: 'Sheet1!A1:B2', style: { bold: true } },
  ], outPath)
  assert.ok(result.patchLog.endsWith('.patch.json'))
  const before = await readWorkbookCells(await readFile(path))
  await rollbackPatchLog(outPath, await readPatchLog(result.patchLog), outPath)
  const restored = await readWorkbookCells(await readFile(outPath))
  assert.deepEqual(restored, before)
})

test('shiftFormulaReferences moves relative rows only when gated by the threshold', () => {
  assert.equal(shiftFormulaReferences('=B2-C2', 'Sheet1', 'Sheet1', { rowDelta: 1, rowThreshold: 3 }), '=B2-C2')
  assert.equal(shiftFormulaReferences('=B3-C3', 'Sheet1', 'Sheet1', { rowDelta: 1, rowThreshold: 3 }), '=B4-C4')
  // Cross-sheet reference stays; same-sheet reference shifts.
  assert.equal(shiftFormulaReferences('=Sheet2!B3-C3', 'Sheet1', 'Sheet1', { rowDelta: 1, rowThreshold: 3 }), '=Sheet2!B3-C4')
})
