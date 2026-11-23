/**
 * Date the `## Unreleased` section and leave a fresh empty one behind.
 *
 * Exported so the part of release.mjs that rewrite history can be tested
 * without a clean tree, a tag, or an actual publish.
 */
export function stampChangelog(changelog, tag, stamp) {
  const dated = changelog.replace(/^## Unreleased[ \t]*$/m, `## ${tag} — ${stamp}`)
  return dated.replace(/^# Changelog[ \t]*$/m, [
    '# Changelog',
    '',
    '## Unreleased',
    '',
    '<!-- Write user-visible changes here before the next release. -->',
  ].join('\n'))
}
