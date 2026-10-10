import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { Context } from '@deepseek-ai/cordis'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import * as plugin from '../src/index.ts'
import { corpusTasks } from '../src/corpus/index.ts'

/**
 * The docs describe the software that exists, not the software that existed.
 *
 * Both kinds of drift here have already happened. `docs/usage.md` told people to pin
 * `dsh-excel-chat@0.23.1` while the package was at 0.42.0, and `docs/benchmark.md`
 * described a 100-task corpus long after it had grown to 144. Neither is a typo; both
 * are a hand-written number outliving the thing it described.
 *
 * The fix is not to be more careful next time. It is to make the claim checkable, so
 * that a release or a corpus change fails the build instead of quietly leaving a
 * sentence that is no longer true.
 */
const ROOT = join(import.meta.dirname, '..')
const manifest = JSON.parse(readFileSync(join(ROOT, 'bundle', 'package.json'), 'utf8')) as {
  name: string
  version: string
}

const DOC_FILES = ['README.md', 'README.en.md', ...readdirSync(join(ROOT, 'docs')).map((name) => join('docs', name))]
  .filter((name) => name.endsWith('.md'))

function docText(name: string): string {
  return readFileSync(join(ROOT, name), 'utf8')
}

test('every pinned install version in the docs is the version that ships', () => {
  // `name@1.2.3` — an instruction a reader will copy. A stale one installs an old
  // release, and the reader has no way to tell.
  const pattern = new RegExp(`${manifest.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}@(\\d+\\.\\d+\\.\\d+)`, 'g')
  const found: Array<{ file: string; version: string }> = []
  for (const file of DOC_FILES) {
    for (const match of docText(file).matchAll(pattern)) {
      found.push({ file, version: match[1]! })
    }
  }
  assert.ok(found.length > 0, 'no pinned install version found in the docs — this guard is not checking anything')
  for (const { file, version } of found) {
    assert.equal(
      version,
      manifest.version,
      `${file} tells readers to install ${manifest.name}@${version}, but ${manifest.version} is what ships`,
    )
  }
})

test('every corpus size stated in the docs is the corpus size', () => {
  // "100 个职场场景" / "100 任务语料" / "100/100 语料回归". Only counts attached to a
  // corpus noun are checked, so an unrelated 100 (a row limit, a token budget) is left
  // alone rather than being forced to mean the corpus.
  const pattern = /(\d+)\s*(?:个)?\s*(?:职场场景|任务语料|文件级任务|语料回归)/g
  const checked: Array<{ file: string; stated: number; line: string }> = []
  for (const file of DOC_FILES) {
    for (const line of docText(file).split('\n')) {
      for (const match of line.matchAll(pattern)) {
        checked.push({ file, stated: Number(match[1]), line: line.trim() })
      }
    }
  }
  assert.ok(checked.length > 0, 'no corpus size found in the docs — this guard is not checking anything')
  for (const { file, stated, line } of checked) {
    assert.equal(
      stated,
      corpusTasks.length,
      `${file} states a corpus of ${stated} but there are ${corpusTasks.length}: ${line}`,
    )
  }
})

test('every tool that registers is documented, in both languages', async () => {
  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(plugin)
  const registered = ctx.tools.schemas().map((schema) => schema.name)
  await ctx.fiber.dispose()

  // Checking the table rather than a count: a count can be right while the table names
  // the wrong tools, and a tool that registers without being documented is the drift
  // that actually costs someone an afternoon.
  for (const file of ['README.md', 'README.en.md']) {
    const text = docText(file)
    const missing = registered.filter((name) => !text.includes(`\`${name}\``))
    assert.deepEqual(missing, [], `${file} does not mention these tools: ${missing.join(', ')}`)
  }
})
