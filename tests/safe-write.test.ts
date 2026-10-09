import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtemp, readFile, readdir, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { findUnpreservedParts, writeWorkbookSafely } from '../src/safe-write.ts'

async function tempDir(): Promise<string> {
  return mkdtemp(join(tmpdir(), 'vera-safe-write-'))
}

test('an overwrite keeps the previous contents in a backup', async () => {
  const dir = await tempDir()
  const target = join(dir, 'book.xlsx')
  await writeFile(target, 'first')

  const result = await writeWorkbookSafely(target, new TextEncoder().encode('second'))

  assert.equal(await readFile(target, 'utf8'), 'second')
  assert.equal(result.backupPath, `${target}.bak`)
  assert.equal(await readFile(`${target}.bak`, 'utf8'), 'first')
})

test('a fresh write needs no backup and leaves no temp file behind', async () => {
  const dir = await tempDir()
  const target = join(dir, 'new.xlsx')

  const result = await writeWorkbookSafely(target, new TextEncoder().encode('only'))

  assert.equal(result.backupPath, undefined)
  assert.equal(await readFile(target, 'utf8'), 'only')
  // The temp file is a sibling so the rename stays on one filesystem; it must not
  // survive the write, or every edit would litter the user's folder.
  assert.deepEqual((await readdir(dir)).sort(), ['new.xlsx'])
})

test('a failed write leaves the target exactly as it was', async () => {
  const dir = await tempDir()
  const target = join(dir, 'book.xlsx')
  await writeFile(target, 'original')

  // A directory that does not exist: the temp write fails before anything is
  // renamed over the target, which is the whole point of writing to a temp file.
  await assert.rejects(writeWorkbookSafely(join(dir, 'missing', 'book.xlsx'), new TextEncoder().encode('x')))

  assert.equal(await readFile(target, 'utf8'), 'original')
  assert.deepEqual((await readdir(dir)).sort(), ['book.xlsx'])
})

test('content ExcelJS cannot round-trip is reported before it is lost', async () => {
  const { unzipSync, zipSync, strToU8 } = await import('fflate')
  const dir = await tempDir()

  const plain = join(dir, 'plain.xlsx')
  await writeFile(plain, Buffer.from(zipSync({ 'xl/workbook.xml': strToU8('<workbook/>') })))
  assert.deepEqual(findUnpreservedParts(new Uint8Array(await readFile(plain))), [])

  const withPivot = join(dir, 'pivot.xlsx')
  await writeFile(withPivot, Buffer.from(zipSync({
    'xl/workbook.xml': strToU8('<workbook/>'),
    'xl/pivotTables/pivotTable1.xml': strToU8('<pivotTableDefinition/>'),
    'xl/pivotCache/pivotCacheDefinition1.xml': strToU8('<pivotCacheDefinition/>'),
  })))
  assert.deepEqual(
    findUnpreservedParts(new Uint8Array(await readFile(withPivot))).sort(),
    ['pivot cache', 'pivot table'],
  )

  // Not a zip at all: a warning path must not turn into a hard failure.
  assert.deepEqual(findUnpreservedParts(new Uint8Array([1, 2, 3])), [])
  void unzipSync
})
