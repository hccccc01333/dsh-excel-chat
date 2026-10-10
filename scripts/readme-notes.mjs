/**
 * Build the "latest release" block the READMEs carry.
 *
 * Kept out of release.mjs so it can be exercised on its own: the release script cannot
 * run in this environment (its `spawnSync git` fails), and prose that goes wrong
 * silently is exactly what this project keeps finding.
 */
import { readFileSync } from 'node:fs'

/** The CHANGELOG section for one tag, or '' when there is none. */
export function changelogSection(changelogPath, tag) {
  const text = readFileSync(changelogPath, 'utf8')
  // The heading carries a date (`## v0.42.0 — 2026-10-10`), so match the tag followed by
  // a space rather than the whole line. Requiring end-of-line here matched nothing, and
  // the failure was silent: an empty block, not an error.
  const after = text.split(new RegExp(`^## ${tag} `, 'm'))[1]
  if (after === undefined) return ''
  return after.split(/^## /m)[0] ?? ''
}

/**
 * The README block for one tag: the version, the date, and where the notes are.
 *
 * It deliberately does not quote the notes. Pulling a summary out of prose was the
 * obvious move and it does not survive contact with the writing: some entries open with
 * a one-line summary, others open mid-sentence and continue, so the block came out with
 * fragments like "…而语料任务的" — and in the English README it quoted Chinese, because
 * the CHANGELOG is written in one language.
 *
 * The release page carries the notes (the workflow builds it from the same section),
 * and GitHub shows the latest release in the repository sidebar. The README's job is to
 * say which version you are looking at and where the detail lives.
 */
export function readmeNotes({ changelogPath, tag, date, heading, linkLabel }) {
  if (changelogSection(changelogPath, tag).trim() === '') return undefined

  return [
    `## ${heading}`,
    '',
    `**${tag}**（${date}）· [${linkLabel}](CHANGELOG.md) · ` +
      `[Releases](https://github.com/hccccc01333/dsh-excel-chat/releases)`,
    '',
  ].join('\n')
}

if (process.argv[1] && process.argv[1].endsWith('readme-notes.mjs')) {
  const [tag, date] = process.argv.slice(2)
  const block = readmeNotes({
    changelogPath: 'CHANGELOG.md',
    tag,
    date,
    heading: '最近更新',
    linkLabel: '完整更新日志',
  })
  if (block === undefined) {
    console.error(`CHANGELOG.md has no '## ${tag}' section`)
    process.exit(1)
  }
  process.stdout.write(block)
}
