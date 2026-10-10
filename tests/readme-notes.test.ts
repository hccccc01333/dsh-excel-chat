import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { changelogSection, readmeNotes } from '../scripts/readme-notes.mjs'

/**
 * The READMEs' "latest release" block is generated at release time.
 *
 * It replaces prose in a file people read first, so getting it wrong is visible and
 * silent at once. The first version matched `^## <tag>$` — but CHANGELOG headings carry
 * a date (`## v0.42.0 — 2026-10-10`), so it matched nothing and would have written an
 * empty block.
 */
const fixture = async (text: string) => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-notes-'))
  const path = join(dir, 'CHANGELOG.md')
  await writeFile(path, text, 'utf8')
  return path
}

test('changelogSection finds a heading that carries a date', async () => {
  const path = await fixture('# Changelog\n\n## v1.2.3 — 2026-01-01\n\n- did a thing\n\n## v1.2.2 — 2025-12-31\n\n- older\n')
  const section = changelogSection(path, 'v1.2.3')
  assert.match(section, /did a thing/)
  assert.ok(!section.includes('older'), 'the section ran past the next heading')
})

test('changelogSection returns empty for a tag that is not there', async () => {
  const path = await fixture('# Changelog\n\n## v1.2.3 — 2026-01-01\n\n- did a thing\n')
  assert.equal(changelogSection(path, 'v9.9.9'), '')
})

test('readmeNotes builds a version line and refuses a missing tag', async () => {
  const path = await fixture('# Changelog\n\n## v1.2.3 — 2026-01-01\n\n- did a thing\n')
  const block = readmeNotes({ changelogPath: path, tag: 'v1.2.3', date: '2026-01-01', heading: '最近更新', linkLabel: '完整更新日志' })
  assert.ok(block)
  assert.match(block, /^## 最近更新\n/)
  assert.match(block, /\*\*v1\.2\.3\*\*（2026-01-01）/)
  assert.match(block, /\[完整更新日志\]\(CHANGELOG\.md\)/)
  // It links to the notes rather than quoting them; see the module for why.
  assert.ok(!block.includes('did a thing'), 'the block quoted the changelog body')

  assert.equal(
    readmeNotes({ changelogPath: path, tag: 'v9.9.9', date: '2026-01-01', heading: '最近更新', linkLabel: '完整更新日志' }),
    undefined,
    'a missing tag should be reported, not written as an empty block',
  )
})

test('the real CHANGELOG yields a block for the published version', async () => {
  const { readFile } = await import('node:fs/promises')
  const { fileURLToPath } = await import('node:url')
  const changelog = fileURLToPath(new URL('../CHANGELOG.md', import.meta.url))
  const published = JSON.parse(await readFile(fileURLToPath(new URL('../bundle/package.json', import.meta.url)), 'utf8')).version
  const block = readmeNotes({ changelogPath: changelog, tag: `v${published}`, date: '2026-01-01', heading: '最近更新', linkLabel: '完整更新日志' })
  assert.ok(block, `CHANGELOG.md has no section for the published version v${published}`)
})
