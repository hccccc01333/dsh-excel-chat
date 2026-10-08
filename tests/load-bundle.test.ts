import { test } from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { Context } from '@deepseek-ai/cordis'
import { CallId } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'

const bundleDir = fileURLToPath(new URL('../bundle', import.meta.url))
const bundleUrl = pathToFileURL(join(bundleDir, 'dist/index.js')).href

test('bundle manifest points at an existing patch and entry', async () => {
  const manifest = JSON.parse(await readFile(join(bundleDir, 'package.json'), 'utf8')) as {
    name: string
    dsh: { bundle: { patch: string } }
  }
  assert.equal(manifest.dsh.bundle.patch, './cordis.patch.yml')
  const patch = await readFile(join(bundleDir, 'cordis.patch.yml'), 'utf8')
  assert.match(patch, new RegExp(manifest.name))
  const entry = await readFile(join(bundleDir, 'dist/index.js'), 'utf8')
  assert.match(entry, /excel_validate_formulas/)
})

test('plugin identity matches the package name (issue #2)', async () => {
  const plugin = await import(bundleUrl)
  assert.equal(plugin.name, 'dsh-excel-chat')
})

test('unloading the plugin unregisters its tools (issue #2 ctx.effect)', async () => {
  const plugin = await import(bundleUrl)
  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  const fiber = await ctx.plugin(plugin)
  assert.ok(ctx.tools.get('excel_operate'), 'tool should be registered while loaded')
  await fiber.dispose()
  assert.equal(ctx.tools.get('excel_operate'), undefined, 'tool should be gone after unload')
})

test('built bundle loads as a plugin and runs a tool', async () => {
  const plugin = await import(bundleUrl)
  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(plugin)
  const result = await ctx.tools.execute({
    signal: new AbortController().signal,
    callId: CallId('vera-bundle-1'),
    name: 'excel_validate_formulas',
    arguments: {
      cells: {
        D2: '=B2-C2',
        D3: '=B3-C3',
        D4: '=B4-C3',
        D5: '=B5-C5',
      },
    },
  })
  assert.equal(result.isError, false)
  const value = result.value as { anomalies: Array<{ cell: string }> }
  assert.ok(value.anomalies.some((anomaly) => anomaly.cell === 'D4'))
})

test('excel_operate runs through the plugin context and re-validates', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('D2').value = { formula: 'B2-C2' }
  sheet.getCell('D3').value = { formula: 'B3-C3' }
  sheet.getCell('D4').value = { formula: 'B4-C4' }
  const dir = await mkdtemp(join(tmpdir(), 'vera-bundle-operate-'))
  const input = join(dir, 'book.xlsx')
  const output = join(dir, 'edited.xlsx')
  await writeFile(input, await workbook.xlsx.writeBuffer())

  const plugin = await import(bundleUrl)
  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(plugin)
  const result = await ctx.tools.execute({
    signal: new AbortController().signal,
    callId: CallId('vera-bundle-operate-1'),
    name: 'excel_operate',
    arguments: {
      path: input,
      operations: [{ op: 'set', cells: { 'Sheet1!D4': '=B4-C3' } }],
      outPath: output,
    },
  })
  assert.equal(result.isError, false)
  const value = result.value as {
    outputPath: string
    validation: { anomalies: Array<{ cell: string }> }
  }
  assert.equal(value.outputPath, output)
  assert.ok(value.validation.anomalies.some((anomaly) => anomaly.cell === 'Sheet1!D4'))
})

test('excel_task names the missing field instead of crashing on a malformed step', async () => {
  // `excel_task` declares its steps as `{ type: 'object', additionalProperties: true }`,
  // so unlike `excel_operate` its operations never meet `excelOperationSchema`.
  // A step missing a nested field therefore reached a handler and surfaced as
  // `TypeError: Cannot read properties of undefined (reading 'toUpperCase')` —
  // a message naming neither the operation nor the field.
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.addRow(['name', 'qty'])
  sheet.addRow(['a', 3])
  const dir = await mkdtemp(join(tmpdir(), 'vera-bundle-task-'))
  const input = join(dir, 'book.xlsx')
  await writeFile(input, await workbook.xlsx.writeBuffer())

  const plugin = await import(bundleUrl)
  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(plugin)
  const result = await ctx.tools.execute({
    signal: new AbortController().signal,
    callId: CallId('vera-bundle-task-1'),
    name: 'excel_task',
    arguments: {
      path: input,
      steps: [{
        name: 'sort',
        operations: [{ op: 'sortRange', range: 'Sheet1!A1:B2', keys: [{ direction: 'desc' }] }],
      }],
    },
  })
  assert.equal(result.isError, true)
  const message = JSON.stringify(result.error ?? result.value)
  assert.match(message, /keys\[0\]\.column 缺失/)
  assert.doesNotMatch(message, /Cannot read properties of undefined/)
})

test('excel_read runs through the plugin context and returns lossless JSON', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = 100
  sheet.getCell('A2').value = { formula: 'A1*2' }
  sheet.getCell('A1').font = { bold: true }
  const dir = await mkdtemp(join(tmpdir(), 'vera-bundle-read-'))
  const input = join(dir, 'book.xlsx')
  await writeFile(input, await workbook.xlsx.writeBuffer())

  const plugin = await import(bundleUrl)
  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(plugin)
  const result = await ctx.tools.execute({
    signal: new AbortController().signal,
    callId: CallId('vera-bundle-read-1'),
    name: 'excel_read',
    arguments: { path: input },
  })
  assert.equal(result.isError, false)
  const value = result.value as { sheets: Array<{ cells: Array<{ id: string; value: number; bold?: boolean }> }> }
  const cell = value.sheets[0]!.cells.find((entry) => entry.id === 'Sheet1!A1')
  assert.equal(cell?.value, 100)
  assert.equal(cell?.bold, true)
})

test('the language config option switches user-facing output', async () => {
  // The loader passes an entry's `config` to `apply` as the second argument
  // (`registry.plugin(plugin, options.config, …)`), so this is the path a real
  // `cordis.patch.yml` override takes. Verified against the built bundle because
  // the wiring — not just the catalog — is what can break.
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('销售')
  sheet.addRow(['区域', '金额'])
  sheet.addRow(['华东', 100])
  sheet.addRow(['华东', null])
  const dir = await mkdtemp(join(tmpdir(), 'vera-bundle-i18n-'))
  const input = join(dir, 'sales.xlsx')
  await writeFile(input, await workbook.xlsx.writeBuffer())

  const plugin = await import(bundleUrl)
  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(plugin, { language: 'en' })
  const result = await ctx.tools.execute({
    signal: new AbortController().signal,
    callId: CallId('vera-bundle-i18n-1'),
    name: 'excel_insight',
    arguments: { path: input },
  })
  assert.equal(result.isError, false)
  const value = result.value as { summary: string; suggestions: string[] }
  assert.match(value.summary, /worksheet\(s\)/)
  assert.ok(value.suggestions.every((suggestion) => !/[\u4e00-\u9fa5]/.test(suggestion)))
  await ctx.fiber.dispose()
})
