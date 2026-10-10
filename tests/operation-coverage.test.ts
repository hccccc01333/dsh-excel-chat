import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFile, readdir } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { excelOperationSchema } from '../src/operation-schema.ts'
import { corpusTasks } from '../src/corpus/index.ts'

/**
 * Every operation the schema allows has to be executed by something.
 *
 * Adding a case to the dispatcher without a test that reaches it is invisible: the
 * code compiles, the schema accepts it, and nothing runs it. A scan found 76 of 77
 * operations executed and `unprotectSheet` never — not a bug that time, but not
 * something anyone could have known without looking either.
 *
 * Counting only the shapes that reach the dispatcher, not names in prose: an operation
 * mentioned in a comment looks covered while nothing calls it. Both the corpus and the
 * `op: 'x'` literals in tests count, because either one runs the operation.
 */
test('every schema operation is executed by the corpus or a test', async () => {
  const declared = excelOperationSchema.oneOf.map((branch) => branch.properties.op.enum[0]!)

  const executed = new Set<string>()
  for (const task of corpusTasks) for (const operation of task.operations ?? []) executed.add(operation.op)

  const dir = fileURLToPath(new URL('.', import.meta.url))
  for (const name of (await readdir(dir)).filter((file) => file.endsWith('.ts'))) {
    const source = await readFile(join(dir, name), 'utf8')
    for (const match of source.matchAll(/\bop:\s*'([A-Za-z]+)'/g)) executed.add(match[1]!)
  }

  const never = declared.filter((operation) => !executed.has(operation)).sort()
  assert.deepEqual(never, [], 'these operations are accepted by the schema but nothing executes them')
})
