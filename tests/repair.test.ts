import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { createLlmRepairAdvisor, type LlmText } from '../src/advisor.ts'
import { generateRepairs, repairWorkbookFile } from '../src/repair.ts'
import { detectTableFromCells } from '../src/tables.ts'
import { validate } from '../src/validator.ts'
import { readWorkbookCells } from '../src/workbook.ts'

const fixturePath = fileURLToPath(new URL('../fixtures/silent-error.xlsx', import.meta.url))
const repairedPath = fileURLToPath(new URL('../fixtures/silent-error.repaired.xlsx', import.meta.url))
const autoTablePath = fileURLToPath(new URL('../fixtures/auto-table.xlsx', import.meta.url))

async function writeFixture(): Promise<void> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sales')
  sheet.getCell('D2').value = { formula: 'B2-C2', result: 40 }
  sheet.getCell('D3').value = { formula: 'B3-C3', result: 80 }
  sheet.getCell('D4').value = { formula: 'B4-C3', result: 150 }
  sheet.getCell('D5').value = { formula: 'B5-C5', result: 160 }
  await mkdir(fileURLToPath(new URL('../fixtures', import.meta.url)), { recursive: true })
  await writeFile(fixturePath, await workbook.xlsx.writeBuffer())
}

test('generateRepairs fixes the silent reference-pattern error', () => {
  const cells = {
    D2: '=B2-C2',
    D3: '=B3-C3',
    D4: '=B4-C3',
    D5: '=B5-C5',
  }
  const repairs = generateRepairs(cells, validate(cells))
  assert.deepEqual(repairs, [{
    id: 'D4',
    kind: 'formula',
    oldValue: '=B4-C3',
    newValue: '=B4-C4',
  }])
})

test('repairWorkbookFile writes a repaired copy that re-validates clean', async () => {
  await writeFixture()
  const repair = await repairWorkbookFile(fixturePath)
  assert.equal(repair.repairs.length, 1)
  assert.equal(repair.before.anomalies.filter((a) => a.kind === 'reference-offset').length, 1)
  assert.equal(repair.after.anomalies.filter((a) => a.kind === 'reference-offset').length, 0)
  const data = await readFile(repairedPath)
  assert.ok(data.length > 0)
})

test('generateRepairs fixes a range tail reference (SUM(B4:C3) -> SUM(B4:C4))', () => {
  const cells = {
    D2: '=SUM(B2:C2)',
    D3: '=SUM(B3:C3)',
    D4: '=SUM(B4:C3)',
    D5: '=SUM(B5:C5)',
  }
  const repairs = generateRepairs(cells, validate(cells))
  assert.deepEqual(repairs, [{
    id: 'D4',
    kind: 'formula',
    oldValue: '=SUM(B4:C3)',
    newValue: '=SUM(B4:C4)',
  }])
})

test('generateRepairs rebuilds both range endpoints when both deviate', () => {
  const cells = {
    D2: '=SUM(B2:C2)',
    D3: '=SUM(B3:C3)',
    D4: '=SUM(B3:C3)',
    D5: '=SUM(B5:C5)',
  }
  const repairs = generateRepairs(cells, validate(cells))
  assert.deepEqual(repairs, [{
    id: 'D4',
    kind: 'formula',
    oldValue: '=SUM(B3:C3)',
    newValue: '=SUM(B4:C4)',
  }])
})

test('generateRepairs keeps absolute modifiers on the untouched range endpoint', () => {
  const cells = {
    D2: '=SUM($B$4:C2)',
    D3: '=SUM($B$4:C3)',
    D4: '=SUM($B$4:C3)',
    D5: '=SUM($B$4:C5)',
  }
  const repairs = generateRepairs(cells, validate(cells))
  assert.deepEqual(repairs, [{
    id: 'D4',
    kind: 'formula',
    oldValue: '=SUM($B$4:C3)',
    newValue: '=SUM($B$4:C4)',
  }])
})

test('auto-detected table schema feeds the LLM repair route', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('B1').value = 'Revenue'
  sheet.getCell('C1').value = 'Cost'
  sheet.getCell('D2').value = { formula: 'B2-C2', result: 40 }
  sheet.getCell('D3').value = { formula: 'SUM(B3:C3)', result: 80 }
  sheet.getCell('D4').value = { formula: 'B4-C4', result: 80 }
  await writeFile(autoTablePath, await workbook.xlsx.writeBuffer())

  const cells = await readWorkbookCells(await readFile(autoTablePath))
  const table = detectTableFromCells(cells)
  assert.deepEqual(table, { sheet: 'Sheet1', columns: { Revenue: 'B', Cost: 'C' } })

  const fakeLlm: LlmText = async () => JSON.stringify({
    repairs: [{
      id: 'D3',
      baseCell: 'D3',
      ir: {
        operation: 'binary',
        left: { kind: 'column', column: 'Revenue' },
        right: { kind: 'column', column: 'Cost' },
        operator: '-',
      },
    }],
  })
  const advisor = createLlmRepairAdvisor(fakeLlm, table)
  const repair = await repairWorkbookFile(autoTablePath, advisor, cells)
  assert.equal(repair.repairs.length, 0)
  assert.equal(repair.llmRepairs.length, 1)
  assert.equal(repair.llmRepairs[0]!.newValue, '=B3-C3')
  assert.equal(repair.after.anomalies.length, 0)
})

test('generateRepairs fills an empty gap by cloning the adjacent formula', () => {
  const cells = {
    D2: '=B2-C2',
    D3: '=B3-C3',
    D5: '=B5-C5',
    D6: '=B6-C6',
  }
  const repairs = generateRepairs(cells, validate(cells))
  assert.deepEqual(repairs, [{
    id: 'D4',
    kind: 'formula',
    oldValue: '',
    newValue: '=B4-C4',
  }])
})

test('empty-gap repair preserves sheet id casing', () => {
  const cells = {
    'Sales!D2': '=B2-C2',
    'Sales!D3': '=B3-C3',
    'Sales!D5': '=B5-C5',
    'Sales!D6': '=B6-C6',
  }
  const repairs = generateRepairs(cells, validate(cells))
  assert.equal(repairs.length, 1)
  assert.equal(repairs[0]!.id, 'Sales!D4')
  assert.equal(repairs[0]!.newValue, '=B4-C4')
})

test('repairWorkbookFile fills an empty gap on disk and re-validates clean', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sales')
  sheet.getCell('D2').value = { formula: 'B2-C2', result: 40 }
  sheet.getCell('D3').value = { formula: 'B3-C3', result: 80 }
  sheet.getCell('D5').value = { formula: 'B5-C5', result: 160 }
  sheet.getCell('D6').value = { formula: 'B6-C6', result: 170 }
  const gapPath = fileURLToPath(new URL('../fixtures/empty-gap.xlsx', import.meta.url))
  await writeFile(gapPath, await workbook.xlsx.writeBuffer())
  const repair = await repairWorkbookFile(gapPath)
  assert.equal(repair.repairs.length, 1)
  assert.equal(repair.repairs[0]!.newValue, '=B4-C4')
  assert.equal(repair.after.anomalies.filter((a) => a.kind === 'empty-gap').length, 0)
})

test('repairWorkbookFile scores the repaired workbook against an oracle', async () => {
  await writeFixture()
  const oracleCells = {
    'Sales!D2': '=B2-C2',
    'Sales!D3': '=B3-C3',
    'Sales!D4': '=B4-C4',
    'Sales!D5': '=B5-C5',
  }
  const repair = await repairWorkbookFile(fixturePath, undefined, undefined, oracleCells)
  assert.equal(repair.oracleScore!.passes, true)
  assert.equal(repair.oracleScore!.accuracy, 1)
})
