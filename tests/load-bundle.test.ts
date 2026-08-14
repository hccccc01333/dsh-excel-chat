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
