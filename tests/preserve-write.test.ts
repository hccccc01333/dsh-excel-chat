import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strToU8, unzipSync, zipSync } from 'fflate'
import ExcelJS from 'exceljs'
import { writeWorkbookPreserving } from '../src/preserve-write.ts'
import { writeWorkbookHealthReport } from '../src/health-report.ts'
import { applyPatchesToWorkbook } from '../src/patch.ts'
import { applyPreserved, capturePreserved } from '../src/preserve.ts'
import { stripPivotTableParts } from '../src/workbook.ts'
import { packageWithPivot } from './pivot-package-fixture.ts'
import { applyOperationsToWorkbook } from '../src/operations.ts'

test('preserving a workbook with two pivot caches keeps both cache relationships', async () => {
  const files = unzipSync(packageWithPivot())
  const book = new TextDecoder().decode(files['xl/workbook.xml'])
    .replace('</pivotCaches>', '<pivotCache cacheId="2" r:id="rId9"/></pivotCaches>')
  files['xl/workbook.xml'] = strToU8(book)
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>',
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition2.xml"/></Relationships>'),
  )
  files['xl/pivotCache/pivotCacheDefinition2.xml'] = files['xl/pivotCache/pivotCacheDefinition1.xml']!
  files['[Content_Types].xml'] = strToU8(
    new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>',
      '<Override PartName="/xl/pivotCache/pivotCacheDefinition2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/></Types>'),
  )
  const dir = await mkdtemp(join(tmpdir(), 'vera-two-caches-'))
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('订单').getCell('A1').value = 'x'
  const path = join(dir, 'input.xlsx')
  await writeFile(path, applyPreserved(
    new Uint8Array(await workbook.xlsx.writeBuffer()),
    capturePreserved(new Uint8Array(zipSync(files))),
  ))
  const out = join(dir, 'out.xlsx')
  await applyOperationsToWorkbook(path, [{ op: 'set', cells: { '订单!A1': 'x' } }], out)
  const result = unzipSync(new Uint8Array(await readFile(out)))
  const rels = new TextDecoder().decode(result['xl/_rels/workbook.xml.rels'])
  assert.equal((rels.match(/Type="[^"]*pivotCacheDefinition"/g) ?? []).length, 2)
  const bookOut = new TextDecoder().decode(result['xl/workbook.xml'])
  const ids = [...bookOut.matchAll(/<pivotCache[^>]*r:id="(rId\d+)"/g)].map((match) => match[1])
  assert.equal(ids.length, 2)
  for (const id of ids) assert.match(rels, new RegExp(`Id="${id}"`))
})

test('creating a pivot next to an existing one keeps both caches in one pivotCaches element', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-pivot-collision-'))
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('A1').value = '地区'
  sheet.getCell('B1').value = '金额'
  sheet.getCell('A2').value = '华东'
  sheet.getCell('B2').value = 5
  const plain = new Uint8Array(await workbook.xlsx.writeBuffer())
  const path = join(dir, 'input.xlsx')
  await writeFile(path, applyPreserved(plain, capturePreserved(packageWithPivot())))
  const outPath = join(dir, 'output.xlsx')
  await applyOperationsToWorkbook(path, [{
    op: 'pivot', source: '订单!A1:B5', rowField: 'A', valueField: 'B',
  }], outPath)
  const files = unzipSync(new Uint8Array(await readFile(outPath)))
  const names = Object.keys(files)
  assert.ok(names.includes('xl/pivotTables/pivotTable1.xml'), 'the original pivot part was overwritten')
  assert.ok(names.includes('xl/pivotTables/pivotTable2.xml'), 'the new pivot part is missing')
  const book = new TextDecoder().decode(files['xl/workbook.xml'])
  assert.equal(book.match(/<pivotCaches\b/g)?.length ?? 0, 1, 'the workbook schema allows one pivotCaches')
  assert.match(book, /cacheId="1"/)
  assert.match(book, /cacheId="2"/)
})

test('writeWorkbookPreserving puts pivot parts back after an ExcelJS-style rewrite', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-preserve-write-'))
  const path = join(dir, 'pivot.xlsx')
  const original = packageWithPivot()
  // What the callers hand ExcelJS: anchors stripped, pivot parts still present
  // in the input but absent from the buffer ExcelJS writes out.
  const input = stripPivotTableParts(original)
  const rewritten = new Uint8Array(zipWithoutPivots(input))
  await writeWorkbookPreserving(original, rewritten, path)

  const files = unzipSync(new Uint8Array(await readFile(path)))
  assert.ok(Object.keys(files).some((name) => name.startsWith('xl/pivotTables/')), 'pivot table part was lost')
  assert.ok(Object.keys(files).some((name) => name.startsWith('xl/pivotCache/')), 'pivot cache was lost')
})

/** A real workbook with pivot parts injected the way operations.ts preserves them. */
async function workbookWithPivot(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-pivot-e2e-'))
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('订单')
  sheet.getCell('D2').value = { formula: 'B2*C2' }
  sheet.getCell('D3').value = { formula: 'B3*C3' }
  sheet.getCell('D4').value = { formula: 'B4*C9' }
  sheet.getCell('D5').value = { formula: 'B5*C5' }
  const plain = new Uint8Array(await workbook.xlsx.writeBuffer())
  const withPivot = applyPreserved(plain, capturePreserved(packageWithPivot()))
  const path = join(dir, 'pivot.xlsx')
  await writeFile(path, withPivot)
  return path
}

function hasPivotParts(data: Uint8Array): boolean {
  const names = Object.keys(unzipSync(data))
  return names.some((name) => name.startsWith('xl/pivotTables/')) && names.some((name) => name.startsWith('xl/pivotCache/'))
}

test('writing a health report keeps the pivot table in the workbook', async () => {
  const path = await workbookWithPivot()
  await writeWorkbookHealthReport(path)
  assert.ok(hasPivotParts(new Uint8Array(await readFile(path))), 'the report stripped the pivot table')
})

test('applying a patch keeps the pivot table in the rewritten workbook', async () => {
  const path = await workbookWithPivot()
  const out = path.replace(/\.xlsx$/, '.patched.xlsx')
  await applyPatchesToWorkbook(path, [{
    id: '订单!D4', kind: 'formula', oldValue: '=B4*C9', newValue: '=B4*C4',
  }], out)
  assert.ok(hasPivotParts(new Uint8Array(await readFile(out))), 'the patch stripped the pivot table')
})

function zipWithoutPivots(data: Uint8Array): Uint8Array {
  // `fflate`'s unzip/zip round trip is enough to model a fresh package. Drop
  // the pivot entries, as ExcelJS does when it writes the workbook back.
  const files = unzipSync(data)
  for (const name of Object.keys(files)) {
    if (/^xl\/(pivotTables|pivotCache)\//.test(name)) delete files[name]
  }
  return zipSync(files)
}
