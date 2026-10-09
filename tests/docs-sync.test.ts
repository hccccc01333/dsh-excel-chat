import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { excelOperationSchema } from '../src/operation-schema.ts'

/**
 * One capability is described in four hand-maintained places.
 *
 * `src/index.ts` registers the tool, `README.md` and `docs/usage.md` each carry a
 * table of them, and the planner prompt lists the operations the model may emit.
 * Nothing connects the four, so adding a tool and forgetting one of the tables is
 * invisible until somebody reads the docs and does not find it.
 *
 * These tests read each list and compare it against the one authoritative source:
 * the registered tools, and the schema. They are cheap and they fail loudly, which
 * is the point — the alternative is discovering the drift in a support question.
 *
 * Extraction note: a table row may name several tools at once
 * (`` | `excel_create_chart` / `excel_modify_chart` | ... ``), so the names are
 * taken from anywhere inside a row rather than only at its start. Matching only
 * row-initial names reported a missing tool that was in fact documented.
 */
const read = async (relative: string) => readFile(fileURLToPath(new URL(relative, import.meta.url)), 'utf8')

/** Every `excel_*` name appearing in a markdown table row. */
function namesInTables(markdown: string): Set<string> {
  const found = new Set<string>()
  for (const row of markdown.matchAll(/^\|.*$/gm)) {
    for (const match of row[0].matchAll(/`(excel_\w+)`/g)) found.add(match[1]!)
  }
  return found
}

async function registeredTools(): Promise<string[]> {
  const source = await read('../src/index.ts')
  return [...source.matchAll(/name: '(excel_\w+)'/g)].map((match) => match[1]!)
}

test('every registered tool appears in the README table', async () => {
  const registered = await registeredTools()
  assert.ok(registered.length >= 20, 'expected the plugin to register its tools')
  const documented = namesInTables(await read('../README.md'))
  const missing = registered.filter((tool) => !documented.has(tool)).sort()
  assert.deepEqual(missing, [], 'these tools are registered but not in the README table')
})

test('every registered tool appears in the usage guide table', async () => {
  const registered = await registeredTools()
  const documented = namesInTables(await read('../docs/usage.md'))
  const missing = registered.filter((tool) => !documented.has(tool)).sort()
  assert.deepEqual(missing, [], 'these tools are registered but not in docs/usage.md')
})

test('the docs do not describe tools that no longer exist', async () => {
  const registered = new Set(await registeredTools())
  for (const [label, file] of [['README.md', '../README.md'], ['docs/usage.md', '../docs/usage.md']] as const) {
    const stale = [...namesInTables(await read(file))].filter((tool) => !registered.has(tool)).sort()
    assert.deepEqual(stale, [], `${label} describes these tools, which are not registered`)
  }
})

test('the planner prompt mentions every operation the schema allows', async () => {
  const prompt = await read('../src/llm-planner.ts')
  const operations = excelOperationSchema.oneOf.map((branch) => branch.properties.op.enum[0]!)
  // Word-boundary, not `"op":"x"`: the catalogue is prose
  // (`insertRows/deleteRows/insertColumns/deleteColumns`), so a JSON-shaped
  // pattern finds barely half of them.
  const missing = operations.filter((op) => !new RegExp(`\\b${op}\\b`).test(prompt)).sort()
  assert.deepEqual(missing, [], 'these operations are reachable but never offered to the planner')
})
