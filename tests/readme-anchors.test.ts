import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

/**
 * In-page links in the READMEs have to point at headings that exist.
 *
 * Both pages now carry a table of contents, and a table of contents with a dead anchor
 * is worse than none: it looks maintained and is not. Renaming a heading is an ordinary
 * edit, and nothing else in the suite would notice that the link to it stopped working.
 *
 * The slug follows GitHub's rules: lowercase, punctuation dropped, spaces to hyphens,
 * CJK kept.
 */
const slug = (heading: string) =>
  heading
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s-]/gu, '')
    .trim()
    .replace(/\s+/g, '-')

for (const name of ['README.md', 'README.en.md']) {
  test(`${name}: every in-page link resolves to a heading`, async () => {
    const path = fileURLToPath(new URL(`../${name}`, import.meta.url))
    const text = await readFile(path, 'utf8')

    const headings = new Set<string>()
    for (const line of text.split('\n')) {
      // Strip the carriage return first. Without `m`, `$` anchors to the end of the
      // string and `.` will not match `\r`, so on a CRLF file the pattern matched no
      // heading at all and every anchor looked broken — which is exactly what happened
      // when the file's line endings changed.
      const match = /^#{1,6}\s+(.*)$/.exec(line.replace(/\r$/, ''))
      if (match) headings.add(slug(match[1]!))
    }

    const links = [...text.matchAll(/\]\(#([^)]+)\)/g)].map((match) => match[1]!)
    assert.ok(links.length > 0, `${name} has no in-page links; did the contents section go away?`)

    const broken = links.filter((link) => !headings.has(link))
    assert.deepEqual(broken, [], `${name} has links pointing at headings that do not exist`)
  })
}
