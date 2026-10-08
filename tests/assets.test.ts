import assert from 'node:assert/strict'
import { test } from 'node:test'
import { existsSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import { collectEvidence, EXPECTED, FEATURES, scenes } from '../scripts/assets.mjs'

/**
 * The promotional images are built from what the engine actually does, so they
 * can quietly become fiction the moment the engine changes — a marketing image
 * that shows a repair the tool no longer performs is worse than no image. This
 * pins the images to the measurement they were built from.
 */
test('the assets still show what the engine actually does', async () => {
  const evidence = await collectEvidence()
  for (const [key, expected] of Object.entries(EXPECTED)) {
    assert.deepEqual(evidence[key], expected, `assets show ${key}=${JSON.stringify(expected)} but the engine produced ${JSON.stringify(evidence[key])}`)
  }
  assert.match(evidence.summary, /修复前 1 个异常，修复后 0 个/)
})

test('every feature card carries a claim and is renderable', async () => {
  const evidence = await collectEvidence()
  assert.ok(FEATURES.length >= 4, 'the grid should cover the whole pitch, not one feature')
  for (const feature of FEATURES) {
    assert.ok(feature.id && feature.title && feature.proof, `feature ${feature.id} is incomplete`)
  }
  // Adding a capability is an entry in FEATURES; the grid must pick it up.
  const grid = scenes(evidence).find((scene) => scene.file === 'feature-grid.png')
  assert.ok(grid, 'no feature grid scene')
  for (const feature of FEATURES) {
    assert.ok(grid.html.includes(feature.title), `the grid does not render ${feature.id}`)
  }
})

test('the READMEs point at assets that exist', async () => {
  const root = fileURLToPath(new URL('..', import.meta.url))
  for (const readme of ['README.md', 'README.en.md']) {
    const text = await import('node:fs/promises').then((fs) => fs.readFile(join(root, readme), 'utf8'))
    for (const match of text.matchAll(/!\[[^\]]*\]\((?!https?:)([^)]+)\)/g)) {
      const file = match[1]
      assert.ok(existsSync(join(root, file)), `${readme} references a missing image: ${file}`)
    }
  }
})
