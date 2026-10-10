import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { Context } from '@deepseek-ai/cordis'
import { CallId } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import * as plugin from '../src/index.ts'
import { buildCorpusWorkbook } from '../src/corpus/helpers.ts'
import { applyOperationsToWorkbook } from '../src/operations.ts'
import { findUnpreservedParts } from '../src/safe-write.ts'

/**
 * Refusing to destroy what the writer cannot carry.
 *
 * ExcelJS models cells and styles; it does not model slicers, macros, embedded objects,
 * custom XML or SmartArt, and those parts simply vanish from anything it writes back.
 * The plugin used to warn and edit anyway — and a warning inside a tool result is not
 * consent. These are real files on disk, because the thing being protected is a file on
 * disk: an in-memory stand-in would not have caught the original bug, which is about
 * what ends up in the user's directory.
 */

/** Add a part to a workbook, the way Excel would when the feature is present. */
async function withPart(source: string, name: string, contents: string): Promise<string> {
  const files = unzipSync(new Uint8Array(await readFile(source)))
  const withPart = { ...files, [name]: strToU8(contents) }
  const target = source.replace(/\.xlsx$/, `-${name.split('/').slice(1, 2)[0]}.xlsx`)
  await writeFile(target, zipSync(withPart))
  return target
}

async function fixture(name: string): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'unpreserved-'))
  return buildCorpusWorkbook(dir, name, [
    { name: '订单', headers: ['产品', '数量'], rows: [['苹果', 10], ['香蕉', 5]] },
  ])
}

test('a plain workbook is unaffected by the gate', async () => {
  const path = await fixture('plain')
  assert.deepEqual(findUnpreservedParts(new Uint8Array(await readFile(path))), [])

  // In-place, and it must simply work — the gate is about features, not about writing.
  const result = await applyOperationsToWorkbook(path, [{ op: 'set', cells: { '订单!B2': '11' } }], path)
  assert.equal(result.warnings.length, 0)
  const after = unzipSync(new Uint8Array(await readFile(path)))
  assert.ok(strFromU8(after['xl/worksheets/sheet1.xml']!).includes('11'))
})

test('an in-place edit is refused when the file holds a part that cannot be carried', async () => {
  const path = await withPart(await fixture('slicer'), 'xl/slicers/slicer1.xml', '<slicers/>')
  const before = await readFile(path)

  await assert.rejects(
    () => applyOperationsToWorkbook(path, [{ op: 'set', cells: { '订单!B2': '11' } }], path),
    (error: Error) => {
      assert.match(error.message, /slicer/)
      return true
    },
  )

  // The point of refusing: the file is byte-for-byte what it was.
  assert.deepEqual(await readFile(path), before, 'a refused edit must not touch the file')
})

test('writing to a different file is allowed, because the original is never touched', async () => {
  const path = await withPart(await fixture('copy'), 'xl/vbaProject.bin', 'macro')
  const before = await readFile(path)
  const out = `${path}.copy.xlsx`

  const result = await applyOperationsToWorkbook(path, [{ op: 'set', cells: { '订单!B2': '11' } }], out)

  assert.deepEqual(await readFile(path), before, 'the original must be byte-identical')
  assert.equal(result.warnings.length, 1)
  assert.match(result.warnings[0]!.message, /VBA macro project/)
  const written = unzipSync(new Uint8Array(await readFile(out)))
  assert.ok(strFromU8(written['xl/worksheets/sheet1.xml']!).includes('11'))
})

test('an in-place edit is allowed when the caller says it accepts the loss', async () => {
  const path = await withPart(await fixture('allow'), 'xl/embeddings/oleObject1.bin', 'ole')

  const result = await applyOperationsToWorkbook(
    path,
    [{ op: 'set', cells: { '订单!B2': '11' } }],
    path,
    { onUnpreserved: 'allow' },
  )

  assert.equal(result.warnings.length, 1)
  // The warning still says what was given up, so the tool result records it.
  assert.match(result.warnings[0]!.message, /embedded object/)
})

test('the same file by another spelling is still the same file', async () => {
  const path = await withPart(await fixture('spelling'), 'xl/slicerCaches/slicerCache1.xml', '<x/>')
  // A path that resolves to the same file must not slip past the gate as "a copy".
  const spelled = join(path, '..', path.split(/[\\/]/).pop()!)

  await assert.rejects(
    () => applyOperationsToWorkbook(path, [{ op: 'set', cells: { '订单!B2': '11' } }], spelled),
    /无法保留/,
  )
})

test('every part the writer cannot carry is detected, and pivots are not', async () => {
  const cases: Array<[string, string]> = [
    ['xl/slicers/slicer1.xml', 'slicer'],
    ['xl/timelines/timeline1.xml', 'timeline'],
    ['xl/vbaProject.bin', 'VBA macro project'],
    ['xl/macrosheets/sheet1.xml', 'macro sheet'],
    ['xl/customXml/item1.xml', 'custom XML'],
    ['xl/ctrlProps/ctrlProp1.xml', 'form control'],
    ['xl/embeddings/oleObject1.bin', 'embedded object'],
    ['xl/diagrams/data1.xml', 'SmartArt diagram'],
  ]
  for (const [name, feature] of cases) {
    const path = await withPart(await fixture(`p-${feature.replace(/\W/g, '')}`), name, 'x')
    assert.deepEqual(
      findUnpreservedParts(new Uint8Array(await readFile(path))),
      [feature],
      `${name} should be reported as ${feature}`,
    )
  }

  // Pivot tables are carried across a rewrite by preserve.ts, so reporting them here
  // would refuse edits that are in fact safe.
  const pivot = await withPart(await fixture('pivot'), 'xl/pivotTables/pivotTable1.xml', '<pivotTableDefinition/>')
  assert.deepEqual(findUnpreservedParts(new Uint8Array(await readFile(pivot))), [])
})

/**
 * The same gate, exercised through the tool the model actually calls.
 *
 * The function-level tests above prove the rule; this proves the rule is reachable.
 * The plugin's own history has a case where a check existed in a function and not on
 * the path the model takes, so the tool call is the one that has to be tested.
 */
test('the gate holds through a real excel_operate call', async () => {
  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(plugin)

  const path = await withPart(await fixture('tool'), 'xl/vbaProject.bin', 'macro')
  const before = await readFile(path)
  const call = (args: Record<string, unknown>) =>
    ctx.tools.execute({
      signal: new AbortController().signal,
      callId: CallId(`gate-${Math.random().toString(36).slice(2)}`),
      name: 'excel_operate',
      arguments: { path, operations: [{ op: 'set', cells: { '订单!B2': '11' } }], ...args },
    })

  // In place: refused, and the file on disk is untouched.
  const refused = await call({ outPath: path })
  assert.equal(refused.isError, true, 'an in-place edit of a macro workbook must be refused')
  assert.deepEqual(await readFile(path), before, 'a refused edit must not touch the file')

  // Explicitly accepted: proceeds, and says what was given up.
  const allowed = await call({ outPath: path, onUnpreserved: 'allow' })
  assert.notEqual(allowed.isError, true, 'an explicit allow must go through')

  await ctx.fiber.dispose()
})
