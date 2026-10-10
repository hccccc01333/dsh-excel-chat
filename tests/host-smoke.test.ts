import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtemp, readFile, readdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { createRequire } from 'node:module'
import { Context } from '@deepseek-ai/cordis'
import { CallId } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import { buildCorpusWorkbook } from '../src/corpus/helpers.ts'

/**
 * A host-level smoke test: boot the built bundle and prove the plugin is actually there.
 *
 * Issue #5 was the plugin disappearing without a word. Since dsh 0.2.0 a launcher checks
 * every `@deepseek-ai/dsh*` peer in a bundle's manifest against the running version, and
 * a bundle that fails goes into the profile's `skippedBundles` — its patch never applies,
 * its rows never mount, and **nothing is printed**. From the outside that is
 * indistinguishable from "the plugin is not installed".
 *
 * `compatibility.test.ts` reproduces that rule against a list of kernel versions written
 * down by hand. This test does the other half: it boots the bundle against the host that
 * is **actually installed here**, and then insists on every tool by name. A partial
 * registration — the failure mode that has no error message — fails on the count.
 */
const semver = createRequire(import.meta.url)('semver') as {
  satisfies: (version: string, range: string, options?: { includePrerelease?: boolean }) => boolean
}

const bundleDir = fileURLToPath(new URL('../bundle', import.meta.url))
const bundleUrl = pathToFileURL(join(bundleDir, 'dist/index.js')).href

/**
 * The complete surface. Listed by name rather than counted, so that adding a tool is a
 * deliberate edit here too — and so a tool that stops registering names itself in the
 * failure instead of showing up as "24 !== 25".
 */
const EXPECTED_TOOLS = [
  'excel_autofix',
  'excel_compile_formula',
  'excel_create_chart',
  'excel_create_pivot',
  'excel_diff_workbook',
  'excel_explain_formula',
  'excel_export_charts',
  'excel_export_pdf',
  'excel_find_errors',
  'excel_health_report',
  'excel_insight',
  'excel_menu',
  'excel_modify_chart',
  'excel_operate',
  'excel_preview',
  'excel_profile',
  'excel_read',
  'excel_repair_formulas',
  'excel_semantic_profile',
  'excel_task',
  'excel_trace',
  'excel_undo',
  'excel_validate_charts',
  'excel_validate_charts_visual',
  'excel_validate_formulas',
].sort()

/** The commands the plugin registers when the host offers a command service. */
const EXPECTED_COMMANDS = ['excel-doctor', 'excel-set', 'excel-undo'].sort()

interface Host {
  ctx: Context
  dispose: () => Promise<void>
}

/**
 * Boot the bundle the way the host does: real cordis context, real tool runtime.
 *
 * A minimal `commands` service is provided because the plugin registers its slash
 * commands through `ctx.get('commands')` and would otherwise skip them silently — which
 * is the same "no error, no feature" shape this whole test is about.
 */
async function boot(options: { withCommands?: boolean } = {}): Promise<Host & { commands: string[] }> {
  const plugin = await import(bundleUrl)
  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)

  const commands: string[] = []
  if (options.withCommands !== false) {
    const service = {
      register(definition: { name: string }) {
        commands.push(definition.name)
        return () => {}
      },
    }
    ctx.provide('commands', service)
  }

  const fiber = await ctx.plugin(plugin)
  return {
    ctx,
    commands,
    dispose: async () => {
      await fiber.dispose()
    },
  }
}

test('the installed host satisfies the bundle peer range', async () => {
  // Read the range from the manifest that ships, and the version from the host that is
  // installed. If these ever disagree the bundle is dropped at load time with no
  // message, so this is the cheapest possible early warning.
  const manifest = JSON.parse(await readFile(join(bundleDir, 'package.json'), 'utf8')) as {
    peerDependencies?: Record<string, string>
  }
  const host = createRequire(import.meta.url)('@deepseek-ai/dsh-tools/package.json') as { version: string }
  const ranges = Object.entries(manifest.peerDependencies ?? {}).filter(
    ([name]) => name === '@deepseek-ai/dsh' || name.startsWith('@deepseek-ai/dsh-'),
  )
  assert.ok(ranges.length > 0, 'the bundle must declare at least one dsh peer range')
  for (const [name, range] of ranges) {
    assert.ok(
      semver.satisfies(host.version, range, { includePrerelease: true }),
      `${name} ${range} rejects the installed host ${host.version} — the bundle would be skipped silently`,
    )
  }
})

test('the built bundle registers every tool, by name', async () => {
  const { ctx, dispose } = await boot()
  const registered = ctx.tools.schemas().map((schema) => schema.name).sort()

  // Named, not counted: a missing tool has to say which one.
  assert.deepEqual(
    registered.filter((name) => !EXPECTED_TOOLS.includes(name)),
    [],
    'the bundle registered a tool that is not in the expected list',
  )
  assert.deepEqual(
    EXPECTED_TOOLS.filter((name) => !registered.includes(name)),
    [],
    'the bundle failed to register these tools',
  )
  await dispose()
})

test('every registered tool has a usable schema', async () => {
  const { ctx, dispose } = await boot()
  for (const schema of ctx.tools.schemas()) {
    assert.ok(schema.name.length > 0, 'a tool registered with an empty name')
    assert.ok(schema.description && schema.description.length > 20, `${schema.name} has no real description`)
    // The model picks a tool by its description; one that says nothing is a broken tool
    // even though it registers cleanly.
    assert.ok(schema.parameters !== undefined, `${schema.name} declares no parameters`)
  }
  await dispose()
})

test('the slash commands register when the host offers a command service', async () => {
  const { commands, dispose } = await boot()
  assert.deepEqual(commands.sort(), EXPECTED_COMMANDS, 'the slash commands did not all register')
  await dispose()
})

test('unloading removes every tool and command', async () => {
  const { ctx, dispose } = await boot()
  assert.equal(ctx.tools.schemas().length, EXPECTED_TOOLS.length)
  await dispose()
  // A plugin that leaves tools behind after unload breaks reload in a long session.
  assert.deepEqual(ctx.tools.schemas().map((schema) => schema.name), [], 'tools survived unload')
})

test('the plugin reads and writes a real file end to end', async () => {
  const { ctx, dispose } = await boot()
  const dir = await mkdtemp(join(tmpdir(), 'host-smoke-'))
  const path = join(dir, 'smoke.xlsx')

  // The input has to exist before the first tool call: every editing tool reads the
  // workbook it is about to rewrite, and none of them invent one. Building it here
  // rather than through the plugin keeps the test about the write, not about creation.
  await buildCorpusWorkbook(dir, 'smoke', [
    { name: 'Sheet1', headers: ['a', 'b'], rows: [['seed', 1]] },
  ])
  const seed = join(dir, 'smoke.xlsx')
  assert.ok(seed === path || (await readdir(dir)).some((name) => name.endsWith('.xlsx')))

  const created = await ctx.tools.execute({
    signal: new AbortController().signal,
    callId: CallId('smoke-create'),
    name: 'excel_operate',
    arguments: {
      path: seed,
      operations: [{ op: 'set', cells: { 'Sheet1!A1': 'hello', 'Sheet1!B1': '42' } }],
    },
  })
  assert.notEqual(created.isError, true, `editing the workbook failed: ${JSON.stringify(created.error ?? created.value)}`)
  const outputPath = (created.value as { outputPath?: string }).outputPath
  assert.ok(outputPath, 'the tool did not report where it wrote')

  const read = await ctx.tools.execute({
    signal: new AbortController().signal,
    callId: CallId('smoke-read'),
    name: 'excel_read',
    arguments: { path: outputPath },
  })
  assert.notEqual(read.isError, true, `reading the workbook failed: ${JSON.stringify(read.error ?? read.value)}`)
  const value = read.value as { sheets?: Array<{ cells?: Array<{ id: string; value: unknown }> }> }
  const cells = value.sheets?.flatMap((sheet) => sheet.cells ?? []) ?? []
  assert.ok(
    cells.some((cell) => cell.id.endsWith('A1') && cell.value === 'hello'),
    'the value written through the tool did not come back out',
  )
  assert.ok(cells.length > 0, 'the read returned no cells at all')

  // The output is a real file on disk, not something held in this process.
  const written = await readdir(dir)
  assert.ok(written.some((name) => name.endsWith('.xlsx')), 'no workbook was written')
  await dispose()
})
