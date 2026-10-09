import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runFileBenchmark, runFileBenchmarkTask } from '../src/file-benchmark.ts'
import { corpusTasks } from '../src/corpus/index.ts'
import { buildDependencyGraph } from '../src/graph.ts'
import { readWorkbookCells } from '../src/workbook.ts'

test('corpus ids are unique', () => {
  const ids = corpusTasks.map((task) => task.id)
  assert.equal(new Set(ids).size, ids.length)
  assert.equal(corpusTasks.length, 100)
})

test('no corpus fixture ships a circular formula', async () => {
  // A fixture whose formula reads its own cell is not a repair task — it is a
  // broken example, and the autofix will happily "fix" it into another circular
  // formula. Two fixtures did exactly that (毛利 written as `=B2-C2` while the
  // formula itself sits in C2), and it stayed invisible until the dependency
  // graph stopped discarding self-edges.
  const offenders: string[] = []
  for (const task of corpusTasks) {
    const dir = await mkdtemp(join(tmpdir(), 'vera-corpus-'))
    const cells = await readWorkbookCells(await readFile(await task.buildInput(dir)))
    const formulas = Object.entries(cells)
      .filter(([, value]) => typeof value === 'string' && value.trim().startsWith('='))
      .map(([id, formula]) => ({ id, formula }))
    const { cycles } = buildDependencyGraph(formulas)
    if (cycles.length > 0) {
      offenders.push(`${task.id}: ${cycles.map((cycle) => cycle.join(' -> ')).join(', ')}`)
    }
  }
  assert.deepEqual(offenders, [])
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

/**
 * An assertion nobody can fail is not an assertion.
 *
 * The corpus only ever asserts that a correct plan passes, so a *lenient* verifier
 * — one that checks fewer conditions than an assertion declares — cannot show up
 * here at all. What can show up is a vacuous assertion: one that keeps passing when
 * the work it describes is not done. These mutations drop the operation each
 * assertion is about and require the task to fail, which is what gives the style
 * assertions (and the conjunctive ones added alongside them) their teeth.
 */
test('dropping the operation a style assertion describes makes the task fail', async () => {
  const { runFileBenchmarkTask } = await import('../src/file-benchmark.ts')
  const styleTasks = ['format-header-bold', 'format-number-format', 'format-fill-color', 'format-wrap-align']

  for (const id of styleTasks) {
    const task = corpusTasks.find((entry) => entry.id === id)!
    assert.ok(task, `missing corpus task ${id}`)

    const intact = await runFileBenchmarkTask(task, await mkdtemp(join(tmpdir(), `vera-mutation-${id}-`)))
    assert.equal(intact.success, true, `${id}: the canonical plan should pass`)

    // Remove the styling the assertion is about; the assertion must notice.
    const stripped = {
      ...task,
      operations: task.operations.filter((operation) => operation.op !== 'style'),
    }
    const mutated = await runFileBenchmarkTask(stripped, await mkdtemp(join(tmpdir(), `vera-mutation-${id}-`)))
    assert.equal(mutated.success, false, `${id}: dropping the style operation still passed — the assertion is vacuous`)
    assert.ok(
      mutated.checksPassed < mutated.checksTotal,
      `${id}: expected at least one assertion to fail after the mutation`,
    )
  }
})

/**
 * A task that asserts nothing must not be able to score a success.
 *
 * `success` used to be `checksPassed === checksTotal`, which a task with no checks
 * satisfies as `0 === 0` — a free pass that quietly lifts the reported rate. The
 * runner now requires `checksTotal > 0`; this makes sure the corpus can never
 * reach the case in the first place, so a typo in a `checks` field is a test
 * failure rather than a better-looking number.
 */
test('every corpus task asserts something', () => {
  const silent = corpusTasks.filter((task) => !task.checks || task.checks.length === 0)
  assert.deepEqual(silent.map((task) => task.id), [], 'these tasks assert nothing and would score a free success')

  for (const task of corpusTasks) {
    for (const check of task.checks) {
      const conditions = ['expect', 'startsWith', 'fill', 'bold', 'numberFormat', 'wrapText', 'hAlign']
        .filter((key) => (check as Record<string, unknown>)[key] !== undefined)
      assert.ok(conditions.length > 0, `${task.id}/${check.id} declares no condition to check`)
    }
  }
})

/**
 * A task must not be scored against a requirement it never states.
 *
 * 17 of the 100 tasks asserted a specific output-sheet name that the description
 * never mentioned — `preset 产品模板：报表+色阶。` was checked against a sheet called
 * `订单-产品分析`. A plan that did the work correctly but named the sheet something
 * else failed, so the success rate charged the agent for not guessing an unstated
 * name. That is not a capability worth measuring. The names are now in the
 * descriptions; this keeps them there.
 */
test('every sheet an assertion names is named in the task description', async () => {
  const ExcelJS = (await import('exceljs')).default
  const dir = await mkdtemp(join(tmpdir(), 'vera-sheet-spec-'))
  const unstated: string[] = []

  for (const task of corpusTasks) {
    const input = await task.buildInput(dir)
    const inputWorkbook = new ExcelJS.Workbook()
    await inputWorkbook.xlsx.load(await readFile(input))
    const inputSheets = new Set(inputWorkbook.worksheets.map((sheet) => sheet.name))

    for (const sheet of new Set(task.checks.map((check) => check.id.split('!')[0]!))) {
      // Sheets that already exist are context; a sheet the plan creates is a
      // requirement, and the description has to say so.
      if (inputSheets.has(sheet)) continue
      if (!task.description.includes(sheet)) unstated.push(`${task.id} asserts sheet 「${sheet}」 but the description never names it`)
    }
  }

  assert.deepEqual(unstated, [])
})
