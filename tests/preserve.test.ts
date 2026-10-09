import assert from 'node:assert/strict'
import { test } from 'node:test'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { applyPreserved, capturePreserved, isEmpty } from '../src/preserve.ts'
import { validateStructure } from '../src/package-check.ts'

/**
 * A minimal package that looks like one Excel wrote around a pivot table: the
 * worksheet relationship is the anchor, and the workbook declares the cache. Built
 * by hand so the test needs no Excel and runs on CI.
 */
function packageWithPivot(): Uint8Array {
  return new Uint8Array(zipSync({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>' +
      '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>' +
      '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="订单" sheetId="1" r:id="rId1"/></sheets>' +
      '<pivotCaches><pivotCache cacheId="1" r:id="rId3"/></pivotCaches></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>' +
      '</Relationships>',
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>',
    ),
    'xl/worksheets/_rels/sheet1.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>' +
      '</Relationships>',
    ),
    'xl/pivotTables/pivotTable1.xml': strToU8('<pivotTableDefinition name="透视表1"><location ref="A3:B7"/></pivotTableDefinition>'),
    'xl/pivotCache/pivotCacheDefinition1.xml': strToU8('<pivotCacheDefinition/>'),
  }))
}

/** What ExcelJS hands back: same workbook, every pivot trace gone. */
function packageWithoutPivot(): Uint8Array {
  return new Uint8Array(zipSync({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="订单" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    // ExcelJS numbers its own relationships, and rId1 is already taken here — the
    // injected cache relationship must not reuse it.
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>',
    ),
    'xl/styles.xml': strToU8('<?xml version="1.0"?><styleSheet/>'),
  }))
}

test('an edit keeps the pivot table parts and every reference to them', () => {
  const preserved = capturePreserved(packageWithPivot())
  assert.deepEqual(Object.keys(preserved.parts).sort(), [
    'xl/pivotCache/pivotCacheDefinition1.xml',
    'xl/pivotTables/pivotTable1.xml',
  ])

  const merged = applyPreserved(packageWithoutPivot(), preserved)
  const files = unzipSync(merged)

  // The parts themselves.
  for (const name of Object.keys(preserved.parts)) assert.ok(files[name], `${name} was not restored`)

  // And the three references Excel needs in order to see them.
  const workbook = strFromU8(files['xl/workbook.xml']!)
  assert.match(workbook, /<pivotCaches>/)
  const cacheId = /<pivotCache[^>]*r:id="([^"]+)"/.exec(workbook)?.[1]
  assert.ok(cacheId, 'the cache declaration lost its r:id')

  const workbookRels = strFromU8(files['xl/_rels/workbook.xml.rels']!)
  assert.match(workbookRels, new RegExp(`Id="${cacheId}"[^>]*pivotCacheDefinition`))

  const sheetRels = strFromU8(files['xl/worksheets/_rels/sheet1.xml.rels']!)
  assert.match(sheetRels, /pivotTable/)

  const contentTypes = strFromU8(files['[Content_Types].xml']!)
  assert.match(contentTypes, /PartName="\/xl\/pivotTables\/pivotTable1\.xml"/)

  assert.deepEqual(validateStructure(merged), [])
})

test('injected relationships get fresh ids instead of colliding with the writer', () => {
  const merged = applyPreserved(packageWithoutPivot(), capturePreserved(packageWithPivot()))
  const files = unzipSync(merged)

  // ExcelJS's output already uses rId1 and rId2 in the workbook rels.
  const rels = strFromU8(files['xl/_rels/workbook.xml.rels']!)
  const ids = [...rels.matchAll(/\bId="([^"]+)"/g)].map((m) => m[1]!)
  assert.equal(new Set(ids).size, ids.length, 'a relationship id was reused')
  assert.match(rels, /Id="rId3"[^>]*pivotCacheDefinition/, 'the cache relationship should take the next free id')

  // The sheet rels were empty, so the first free id is rId1 — and the reference
  // in workbook.xml has to agree with whatever was chosen.
  const cacheId = /<pivotCache[^>]*r:id="([^"]+)"/.exec(strFromU8(files['xl/workbook.xml']!))?.[1]
  assert.equal(cacheId, 'rId3')
})

test('a package with no pivot content is passed through untouched', () => {
  const plain = packageWithoutPivot()
  const preserved = capturePreserved(plain)
  assert.equal(isEmpty(preserved), true)
  assert.deepEqual([...applyPreserved(plain, preserved)], [...plain])
})

test('the structural check rejects references that do not resolve', () => {
  assert.deepEqual(validateStructure(packageWithPivot()), [])

  // A relationship pointing at a part that is not there: exactly the mistake
  // hand-written package surgery makes, and the reason a write is refused.
  const broken = new Uint8Array(zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>'),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="x" Target="worksheets/missing.xml"/></Relationships>',
    ),
    'xl/workbook.xml': strToU8('<?xml version="1.0"?><workbook/>'),
  }))
  const problems = validateStructure(broken)
  assert.ok(problems.some((problem) => /missing\.xml/.test(problem)), problems.join('; '))
})

test('the structural check notices pivot parts Excel could not reach', () => {
  // The parts are in the package but nothing declares the cache, so Excel shows no
  // pivot table — indistinguishable from having lost it.
  const orphaned = new Uint8Array(zipSync({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="x"/></Types>',
    ),
    'xl/workbook.xml': strToU8('<?xml version="1.0"?><workbook/>'),
    'xl/_rels/workbook.xml.rels': strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'),
    'xl/pivotTables/pivotTable1.xml': strToU8('<pivotTableDefinition/>'),
  }))
  const problems = validateStructure(orphaned)
  assert.ok(problems.some((problem) => /pivotCacheDefinition/.test(problem)), problems.join('; '))
  assert.ok(problems.some((problem) => /pivotCaches/.test(problem)), problems.join('; '))
})

/**
 * The synthetic packages above prove the merge logic; this proves it against a
 * file Excel actually wrote, where the part layout, the relationship numbering and
 * the surrounding parts (printerSettings, theme, styles) are all real. The fixture
 * was produced by Excel's own pivot-table API and is checked in, so this runs on CI
 * where Excel does not exist.
 */
test('editing a real Excel pivot-table workbook keeps the pivot table', async () => {
  const { readFile, mkdtemp } = await import('node:fs/promises')
  const { tmpdir } = await import('node:os')
  const { join } = await import('node:path')
  const { fileURLToPath } = await import('node:url')
  const { applyOperationsToWorkbook } = await import('../src/operations.ts')

  const fixture = fileURLToPath(new URL('./fixtures/pivot-table.xlsx', import.meta.url))
  const dir = await mkdtemp(join(tmpdir(), 'vera-pivot-'))
  const out = join(dir, 'edited.xlsx')

  const original = unzipSync(new Uint8Array(await readFile(fixture)))
  const pivotParts = Object.keys(original).filter((name) => /pivot/i.test(name))
  assert.equal(pivotParts.length, 5, 'the fixture should carry a real pivot table')

  const result = await applyOperationsToWorkbook(fixture, [{ op: 'set', cells: { '订单!E1': '备注' } }], out)

  const edited = unzipSync(new Uint8Array(await readFile(out)))
  for (const name of pivotParts) {
    assert.ok(edited[name], name + ' did not survive the edit')
    // Copied, not regenerated — a rebuilt pivot part would not be the user's.
    assert.deepEqual([...edited[name]!], [...original[name]!], name + ' was modified, not copied')
  }

  // The edit still has to have happened.
  assert.match(strFromU8(edited['xl/sharedStrings.xml']!), /备注/)

  // And the package has to be one Excel can open.
  assert.deepEqual(validateStructure(new Uint8Array(await readFile(out))), [])

  // Pivot tables are preserved now, so they must not be reported as lost.
  const warned = result.warnings.filter((warning) => /pivot/i.test(warning.message))
  assert.equal(warned.length, 0, 'pivot tables are preserved but were warned about')
})
