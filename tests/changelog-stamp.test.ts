import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFileSync } from 'node:fs'
import { stampChangelog } from '../scripts/changelog-stamp.mjs'

const ROOT = new URL('..', import.meta.url)

test('stampChangelog dates the notes and restores an empty Unreleased section', () => {
  const changelog = ['# Changelog', '', '## Unreleased', '', '- 一条改动', '', '## v1.0.0', '', '- 旧改动', ''].join('\n')
  const stamped = stampChangelog(changelog, 'v1.1.0', '2026-11-23')
  assert.match(stamped, /^# Changelog\n\n## Unreleased\n\n<!--/)
  assert.match(stamped, /^## v1\.1\.0 — 2026-11-23\n\n- 一条改动/m)
  assert.match(stamped, /^## v1\.0\.0\n/m)
  // The fresh guard section is comment-only, so the next release still fails
  // loudly if nobody writes notes under it.
  const body = stamped.split(/^## Unreleased[ \t]*$/m)[1]!.split(/^## /m)[0]
  assert.equal(body.replace(/<!--[\s\S]*?-->/g, '').trim(), '')
})

test('the checked-in changelog has a non-empty Unreleased section', () => {
  const changelog = readFileSync(new URL('CHANGELOG.md', ROOT), 'utf8')
  assert.match(changelog, /^# Changelog\n/m)
  assert.match(changelog, /^## Unreleased[ \t]*$/m)
  const body = changelog.split(/^## Unreleased[ \t]*$/m)[1]!.split(/^## /m)[0]
  assert.ok(body.replace(/<!--[\s\S]*?-->/g, '').trim().length > 0)
})
