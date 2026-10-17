import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { join } from 'node:path'
import { test } from 'node:test'

const semver = createRequire(import.meta.url)('semver') as {
  satisfies: (version: string, range: string, options?: { includePrerelease?: boolean }) => boolean
}

/**
 * Since dsh 0.2.0 a launcher checks every `@deepseek-ai/dsh*` peer in a bundle's
 * package.json against the running dsh version, and a bundle that fails goes into
 * the profile's `skippedBundles`: its patch layer never applies, its rows never
 * mount, and **nothing is printed**. The plugin then looks like it was never
 * installed — no tools, no error — which is exactly what issue #5 reported, and
 * the only reason 0.38.1 kept working there was a profile exemption pinned to
 * that exact version.
 *
 * This reproduces the host's own rule (`evaluatePluginCompatibility` in
 * `@deepseek-ai/dsh-app-boot`, whose peers are checked with
 * `semver.satisfies(runtime, range, { includePrerelease: true })`) so a peer range
 * that would silently drop the bundle fails here instead.
 */
function incompatibleDshPeers(manifest: { peerDependencies?: Record<string, string> }, runtime: string) {
  const incompatible: Record<string, string> = {}
  for (const [name, range] of Object.entries(manifest.peerDependencies ?? {})) {
    if (name !== '@deepseek-ai/dsh' && !name.startsWith('@deepseek-ai/dsh-')) continue
    if (range.trim() === '' || !semver.satisfies(runtime, range, { includePrerelease: true })) {
      incompatible[name] = range
    }
  }
  return incompatible
}

const manifest = JSON.parse(
  readFileSync(join(import.meta.dirname, '..', 'bundle', 'package.json'), 'utf8'),
) as { peerDependencies?: Record<string, string> }

/** dsh kernels this bundle must not be silently dropped by. */
const SUPPORTED_KERNELS = [
  '0.1.0-rc.5',
  '0.1.0-rc.6',
  '0.1.0-rc.8',
  '0.1.7-rc.2',
  '0.2.0-rc.1',
  '0.2.0-rc.2',
  '0.2.1-alpha.1',
  '0.2.1-alpha.5',
  '0.2.1-rc.2',
]

test('every published dsh kernel satisfies the bundle peer range', () => {
  const rejected = SUPPORTED_KERNELS.filter(
    (runtime) => Object.keys(incompatibleDshPeers(manifest, runtime)).length > 0,
  )
  assert.deepEqual(
    rejected,
    [],
    `dsh would silently skip this bundle on these kernels: ${rejected.join(', ')}`,
  )
})

test('the peer range still stops short of the next kernel minor', () => {
  // Claiming a kernel we have never run on would trade a silent skip for a crash.
  //
  // `0.3.0-rc.1` is checked as well as `0.3.0`, because `<0.3.0` alone does not exclude it:
  // the host evaluates peers with `includePrerelease`, and a prerelease sorts *below* its
  // release, so a plain upper bound of `0.3.0` admits every `0.3.0-rc.*`. The range ends at
  // `<0.3.0-0` — the lowest possible 0.3.0 prerelease — which is what actually stops them.
  for (const untested of ['0.3.0-rc.1', '0.3.0']) {
    const peers = incompatibleDshPeers(manifest, untested)
    assert.ok(
      Object.keys(peers).length > 0,
      `the range claims dsh ${untested}, which this plugin has never been tested against`,
    )
  }
})

test('the check rejects the range that caused issue #5', () => {
  // The published 0.38.1 and 0.39.0 manifests both declared `^0.1.0-rc.6`, which
  // dsh 0.2.0-rc.2 does not satisfy — so this assertion is what proves the check
  // above can actually fail, rather than passing vacuously.
  const broken = { peerDependencies: { '@deepseek-ai/dsh-tools': '^0.1.0-rc.6' } }
  const peers = incompatibleDshPeers(broken, '0.2.0-rc.2')
  assert.deepEqual(Object.keys(peers), ['@deepseek-ai/dsh-tools'])
})

test('the bundle declares the dsh peers the host checks', () => {
  const checked = Object.keys(manifest.peerDependencies ?? {}).filter(
    (name) => name === '@deepseek-ai/dsh' || name.startsWith('@deepseek-ai/dsh-'),
  )
  assert.deepEqual(checked.sort(), [
    '@deepseek-ai/dsh-attachment',
    '@deepseek-ai/dsh-llm',
    '@deepseek-ai/dsh-system-prompt',
    '@deepseek-ai/dsh-tools',
  ])
})
