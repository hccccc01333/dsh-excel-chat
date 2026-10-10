#!/usr/bin/env node
/**
 * Cut a release in one command.
 *
 *   node scripts/release.mjs 0.39.6
 *   node scripts/release.mjs 0.39.6 --dry-run     # stop before writing anything
 *   node scripts/release.mjs 0.39.6 --no-push     # commit and tag locally only
 *
 * It bumps `bundle/package.json` (the version npm and the publish workflow
 * check) and the root workspace manifest, stamps the CHANGELOG's `## Unreleased`
 * section with the version and today's date, rebuilds `bundle/dist`, runs the
 * test suite, commits, tags and pushes. The tag push is what triggers
 * `.github/workflows/publish.yml`, which publishes to npm over OIDC — no token
 * and no 2FA prompt — and opens the GitHub Release.
 *
 * The tree must be clean and the CHANGELOG must carry an `## Unreleased`
 * section: the notes are written by hand, and this script only dates them.
 */
import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { readmeNotes } from './readme-notes.mjs'

const ROOT = dirname(dirname(fileURLToPath(import.meta.url)))
const BUNDLE_MANIFEST = join(ROOT, 'bundle', 'package.json')
const ROOT_MANIFEST = join(ROOT, 'package.json')
const CHANGELOG = join(ROOT, 'CHANGELOG.md')

const args = process.argv.slice(2)
const flags = new Set(args.filter((arg) => arg.startsWith('--')))
const version = args.find((arg) => !arg.startsWith('--'))
const dryRun = flags.has('--dry-run')
const push = !flags.has('--no-push')

function fail(message) {
  console.error(`release: ${message}`)
  process.exit(1)
}

function run(command, argv, options = {}) {
  return execFileSync(command, argv, { cwd: ROOT, encoding: 'utf8', stdio: options.quiet ? 'pipe' : 'inherit' })
}

if (version === undefined) fail('usage: node scripts/release.mjs <version> [--dry-run] [--no-push]')
if (!/^\d+\.\d+\.\d+$/.test(version)) fail(`version ${JSON.stringify(version)} is not X.Y.Z`)

// A release is a promise that this exact tree shipped; refuse to guess.
const dirty = run('git', ['status', '--porcelain'], { quiet: true }).trim()
if (dirty !== '') fail(`the working tree has uncommitted changes:\n${dirty}`)

const branch = run('git', ['rev-parse', '--abbrev-ref', 'HEAD'], { quiet: true }).trim()
if (branch !== 'master') fail(`expected to release from master, currently on ${branch}`)

const tag = `v${version}`
const existing = run('git', ['tag', '--list', tag], { quiet: true }).trim()
if (existing !== '') fail(`${tag} already exists locally`)
const remote = run('git', ['ls-remote', '--tags', 'origin', `refs/tags/${tag}`], { quiet: true }).trim()
if (remote !== '') fail(`${tag} already exists on origin`)

const changelog = readFileSync(CHANGELOG, 'utf8')
if (!/^## Unreleased[ \t]*$/m.test(changelog)) {
  fail('CHANGELOG.md has no `## Unreleased` section; write the release notes under it first')
}
const unreleasedBody = changelog
  .split(/^## Unreleased[ \t]*$/m)[1]
  .split(/^## /m)[0]
  // The section ships with a guiding comment; it is not release notes.
  .replace(/<!--[\s\S]*?-->/g, '')
  .trim()
if (unreleasedBody === '') fail('the `## Unreleased` section is empty; there is nothing to release')

const current = JSON.parse(readFileSync(BUNDLE_MANIFEST, 'utf8')).version
console.log(`release: ${current} -> ${version}  (tag ${tag}, branch ${branch})`)
if (dryRun) {
  console.log('release: --dry-run, stopping before any write')
  process.exit(0)
}

// 1. Manifests. `npm version` keeps package-lock in step with the root.
writeFileSync(BUNDLE_MANIFEST, readFileSync(BUNDLE_MANIFEST, 'utf8').replace(
  `"version": "${current}"`,
  `"version": "${version}"`,
))
run('npm', ['version', version, '--no-git-tag-version', '--allow-same-version'], { quiet: true })

// 2. Date the notes. The body is untouched: this script does not write prose.
const today = new Date()
const stamp = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`
writeFileSync(CHANGELOG, changelog.replace(/^## Unreleased[ \t]*$/m, `## ${tag} — ${stamp}`))

// 2b. Refresh the "latest release" block in both READMEs.
//
// A hand-written block goes stale the moment the next version ships and nothing would
// notice, so it is generated. The builder lives in `readme-notes.mjs` rather than here
// because this script cannot run in every environment (its `spawnSync git` fails in the
// sandbox), and something that silently writes the wrong prose should be testable.
for (const [file, heading, linkLabel] of [
  ['README.md', '最近更新', '完整更新日志'],
  ['README.en.md', "What's new", 'Full changelog'],
]) {
  const block = readmeNotes({ changelogPath: CHANGELOG, tag, date: stamp, heading, linkLabel })
  if (block === undefined) fail(`CHANGELOG.md has no '## ${tag}' section to summarise`)
  const text = readFileSync(file, 'utf8')
  const pattern = new RegExp(`^## ${heading}[ \\t]*$[\\s\\S]*?(?=^## )`, 'm')
  if (!pattern.test(text)) fail(`${file} has no '## ${heading}' section to refresh`)
  writeFileSync(file, text.replace(pattern, block))
}

// 3. The published bundle must match the tag, so rebuild before testing.
run('npm', ['run', 'build:bundle'])
run('node', ['--test', 'tests/*.test.ts'])

// 4. Commit, tag, push. The tag push is what starts the publish workflow.
run('git', ['add', '--all'])
run('git', ['commit', '-m', `chore(release): ${tag}`])
run('git', ['tag', '-a', tag, '-m', tag])
if (push) {
  run('git', ['push', 'origin', branch, '--no-follow-tags'])
  run('git', ['push', 'origin', tag])
  console.log(`release: pushed ${tag}; .github/workflows/publish.yml will publish ${version}`)
} else {
  console.log(`release: committed and tagged locally. Push with:`)
  console.log(`  git push origin ${branch} --no-follow-tags && git push origin ${tag}`)
}
