/**
 * The version the docs tell readers to install.
 *
 * `docs-consistency.test.ts` refuses to let a doc pin a version other than the one that
 * ships — a stale `dsh-excel-chat@0.23.1` sends a reader to a release from months ago and
 * nothing about the sentence looks wrong. That guard is the reason this module exists:
 * `release.mjs` bumps the manifests, and if nothing bumps the docs in the same breath the
 * guard fails inside the release script's own test step and the release cannot finish.
 *
 * Kept out of `release.mjs` for the same reason `readme-notes.mjs` is: the release script
 * cannot run in every environment (its `spawnSync git` fails in the sandbox), so anything
 * it writes has to be exercisable on its own. The file list and the pattern live here
 * rather than in both places because two copies of one rule drift — the guard would check
 * a file the release step never updates, and the failure would look like a mystery.
 */
import { readdirSync } from 'node:fs'
import { join } from 'node:path'

/** Every doc that may pin an install version, relative to the repository root. */
export function docFiles(root) {
  return ['README.md', 'README.en.md', ...readdirSync(join(root, 'docs')).map((name) => join('docs', name))]
    .filter((name) => name.endsWith('.md'))
}

/** `name@1.2.3` — an instruction a reader will copy. */
export function pinPattern(packageName) {
  return new RegExp(`${packageName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}@(\\d+\\.\\d+\\.\\d+)`, 'g')
}

/** The versions one document pins, in the order they appear. */
export function findPinnedVersions(text, packageName) {
  return [...text.matchAll(pinPattern(packageName))].map((match) => match[1])
}

/**
 * The same document with every pinned version replaced by the one that ships.
 *
 * The replacement is a function, not a string: `$&` and friends are meaningful in a string
 * replacement, and a package name is not the place to find that out.
 */
export function stampPinnedVersions(text, packageName, version) {
  return text.replace(pinPattern(packageName), () => `${packageName}@${version}`)
}
