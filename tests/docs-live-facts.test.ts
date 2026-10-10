import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

/**
 * Live documents must not state a test count.
 *
 * The count was written by hand in the README and in CONTRIBUTING, and it was wrong three
 * times in one session — 423, then 483, then 489, each stale the moment tests were added.
 * The command prints the number itself, so writing it down buys nothing and costs a
 * correction every time.
 *
 * CHANGELOG.md is exempt and deliberately so: "483 → 489 tests" there is a historical
 * record of what a release shipped, which stays true. The rule is about live counts, not
 * about the string.
 */
const LIVE_DOCS = ['README.md', 'README.en.md']

/**
 * "（483 项）", "483 tests", "483 项测试" — a count near the word for tests, or the
 * parenthesised-count shape that reads as one in a test command.
 *
 * The first version only matched a count *adjacent to* the word, and the verification
 * injection `# 全量（502 项）` slipped straight through it — a check that could not fail
 * is not a check. The bracketed form is now matched on its own.
 */
const COUNT_PATTERNS = [
  /\d+\s*项\s*测试/,
  /测试\s*\d+\s*项/,
  /\d+\s+tests?\b/i,
  /\btests?\s*[:：]\s*\d+/i,
  /[（(]\s*\d+\s*项\s*[）)]/,
  /[（(]\s*\d+\s+tests?\s*[）)]/i,
]

const read = (name: string) => readFile(fileURLToPath(new URL(`../${name}`, import.meta.url)), 'utf8')

for (const name of LIVE_DOCS) {
  test(`${name} does not hard-code a test count`, async () => {
    const text = await read(name)
    for (const pattern of COUNT_PATTERNS) {
      const match = pattern.exec(text)
      assert.equal(
        match,
        null,
        `${name} states a test count ("${match?.[0]}"); the command prints it, and a written one goes stale`,
      )
    }
  })
}

