import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFile, readdir, stat } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

/**
 * User-facing warning text has to go through `t()`.
 *
 * The coverage test in `i18n.test.ts` counts `t()` keys, so a message written outside
 * `t()` is invisible to it: `message: 'uniqueValues extracted N value(s)'` reads as
 * covered while staying English in Chinese mode — and Chinese is the default. A scan
 * found 37 of them, in the operation warnings a user sees most.
 *
 * This checks the shape rather than the wording: a `message:` whose value contains a
 * long literal and no `t(` call is text somebody will read in the wrong language.
 */
async function sourceFiles(dir: string): Promise<string[]> {
  const found: string[] = []
  for (const entry of await readdir(dir)) {
    const path = join(dir, entry)
    if ((await stat(path)).isDirectory()) found.push(...(await sourceFiles(path)))
    else if (path.endsWith('.ts')) found.push(path)
  }
  return found
}

test('warning messages are passed through t()', async () => {
  const root = fileURLToPath(new URL('../src', import.meta.url))
  const offenders: string[] = []
  for (const file of await sourceFiles(root)) {
    const source = await readFile(file, 'utf8')
    for (const match of source.matchAll(/message:\s*(.+)/g)) {
      const expression = match[1]!.trim()
      if (/\bt\(/.test(expression)) continue
      // Long literal text: prose somebody reads, as opposed to a short identifier.
      if (!/['"`][A-Za-z][^'"`]{12,}['"`]/.test(expression)) continue
      offenders.push(`${file.replace(root, 'src')}: ${expression.slice(0, 60)}`)
    }
  }
  assert.deepEqual(offenders, [], 'these warnings never reach the translation table')
})
