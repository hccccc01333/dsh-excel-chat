import { test } from 'node:test'
import assert from 'node:assert/strict'
import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import { createPivotTable } from '../src/pivot.ts'
import { makeSalesWorkbook } from './sales-fixture.ts'

const excelAvailable = existsSync('C:/Program Files/Microsoft Office/Root/Office16/EXCEL.EXE')

test('createPivotTable builds a native pivot via Excel COM', { skip: !excelAvailable, timeout: 180000 }, async () => {
  const source = await makeSalesWorkbook()
  const outPath = join(join(source, '..'), 'pivot.xlsx')
  const result = await createPivotTable(source, {
    sheet: '订单',
    range: '订单!A1:F7',
    rows: ['B'],
    values: [
      { column: 'F', function: 'sum' },
      { column: 'D', function: 'count' },
    ],
    outputSheet: '透视表',
  }, outPath)
  assert.equal(result.pivotSheet, '透视表')
  assert.equal(result.groups, 3)
  assert.equal(result.recordCount, 6)

  const files = unzipSync(new Uint8Array(await readFile(outPath)))
  const keys = Object.keys(files)
  assert.ok(keys.some((key) => key.startsWith('xl/pivotTables/')), 'missing pivot table part')
  assert.ok(keys.some((key) => key.startsWith('xl/pivotCache/')), 'missing pivot cache parts')
  const workbookXml = strFromU8(files['xl/workbook.xml'] ?? new Uint8Array(0))
  assert.ok(workbookXml.includes('pivotCaches'))
  const pivotTableXml = strFromU8(
    files[keys.find((key) => key.startsWith('xl/pivotTables/') && key.endsWith('.xml'))!] ?? new Uint8Array(0),
  )
  assert.ok(pivotTableXml.includes('金额') || pivotTableXml.includes('Sum of 金额'))
})
