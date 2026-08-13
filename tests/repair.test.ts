import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { generateRepairs, repairWorkbookFile } from '../src/repair.ts'
import { validate } from '../src/validator.ts'

const fixturePath = fileURLToPath(new URL('../fixtures/silent-error.xlsx', import.meta.url))
const repairedPath = fileURLToPath(new URL('../fixtures/silent-error.repaired.xlsx', import.meta.url))

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
