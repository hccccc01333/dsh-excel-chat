import { test } from 'node:test'
import assert from 'node:assert/strict'
import { runFileBenchmark } from '../src/file-benchmark.ts'
import { corpusTasks } from '../src/corpus/index.ts'

test('corpus ids are unique', () => {
  const ids = corpusTasks.map((task) => task.id)
  assert.equal(new Set(ids).size, ids.length)
  assert.equal(corpusTasks.length, 38)
})

test('every canonical plan passes its checks with clean workbook integrity', async () => {
  const report = await runFileBenchmark(corpusTasks)
  assert.equal(report.total, corpusTasks.length)
  assert.equal(report.success, corpusTasks.length)
  assert.equal(report.successRate, 1)
  assert.equal(report.integrityRate, 1)
  assert.equal(report.meanAccuracy, 1)
  for (const category of ['editing', 'analysis', 'formula', 'workflow']) {
    const entry = report.categories[category]
    assert.ok(entry, `missing category ${category}`)
    assert.ok(entry.total > 0)
    assert.equal(entry.success, entry.total, `category ${category} not fully green`)
  }
})
