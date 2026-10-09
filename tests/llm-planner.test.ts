import { test } from 'node:test'
import assert from 'node:assert/strict'
import { createLlmPlanner } from '../src/llm-planner.ts'
import { excelOperationSchema } from '../src/operation-schema.ts'

test('planner catalog covers every excel_operate operation', async () => {
  let captured = ''
  const planner = createLlmPlanner(async (prompt) => {
    captured = prompt
    return JSON.stringify({ steps: [{ name: 's', operations: [] }] })
  })
  await planner.plan({
    goal: '随便做点什么',
    path: 'D:/sales.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '',
    validationSummary: '',
  })
  // Derived from the schema, so adding an operation without teaching the
  // planner about it fails here instead of silently becoming unreachable.
  const operations: string[] = excelOperationSchema.oneOf.map(
    (entry: any) => entry.properties.op.enum[0],
  )
  const missing = operations.filter((op) => !new RegExp(`\\b${op}\\b`).test(captured))
  assert.deepEqual(
    missing,
    [],
    `in the schema but absent from the planner catalog, so the model can never plan them: ${missing.join(', ')}`,
  )
})

test('verifier prompt requires evidence-based checklist', async () => {
  let captured = ''
  const planner = createLlmPlanner(async (prompt) => {
    captured = prompt
    return JSON.stringify({ achieved: true, reason: 'ok' })
  })
  await planner.verify({
    goal: '按区域汇总金额',
    path: 'D:/sales.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '订单：3 行 × 4 列',
    validationSummary: '0 个公式异常',
    previousPlan: undefined,
    previousResult: undefined,
    verifierNote: undefined,
    executedPlan: [{ name: '汇总', operations: [] }],
    executedResult: { outputPath: 'D:/out.xlsx', steps: [], finalAnomalies: 0 },
    cellSnapshot: '订单!A1=区域',
  })
  assert.match(captured, /可检查点/)
  assert.match(captured, /每一条都要有明确证据/)
})

test('verifier prompt declares snapshot sampling and per-goal evidence standards', async () => {
  let captured = ''
  const planner = createLlmPlanner(async (prompt) => {
    captured = prompt
    return JSON.stringify({ achieved: false, reason: 'no' })
  })
  await planner.verify({
    goal: '做交叉表',
    path: 'D:/sales.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '订单：3 行 × 3 列',
    validationSummary: '0 个公式异常',
    previousPlan: undefined,
    previousResult: undefined,
    verifierNote: undefined,
    executedPlan: [{ name: 'x', operations: [] }],
    executedResult: { outputPath: 'D:/out.xlsx', steps: [], finalAnomalies: 0 },
    cellSnapshot: '订单!A1=区域',
  })
  assert.match(captured, /按工作表轮询采样/)
  assert.match(captured, /SUMIFS/)
  assert.match(captured, /反例/)
})

test('planner prompt includes new ops, few-shot examples, and analysis rules', async () => {
  let captured = ''
  const planner = createLlmPlanner(async (prompt) => {
    captured = prompt
    return JSON.stringify({ steps: [{ name: 's', operations: [{ op: 'aggregateReport' }] }] })
  })
  const plan = await planner.plan({
    goal: '按区域汇总金额',
    path: 'D:/sales.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '订单：3 行 × 3 列',
    validationSummary: '0 个公式异常',
  })
  const steps = Array.isArray(plan) ? plan : plan.steps
  assert.match(captured, /crosstab/)
  assert.match(captured, /joinSheets/)
  assert.match(captured, /rankColumn/)
  assert.match(captured, /完整示例/)
  assert.match(captured, /禁止用 set 写死汇总数字/)
  assert.match(captured, /metric 是对象/)
  assert.match(captured, /assertions/)
  assert.match(captured, /机器断言/)
})

test('planner normalizes crosstab flat metric and single-string arrays', async () => {
  const planner = createLlmPlanner(async () => JSON.stringify({
    steps: [{
      name: 's',
      operations: [
        { op: 'crosstab', source: '订单!A1:C4', rowColumn: 'A', columnColumn: 'B', metricColumn: 'C', metricFunction: 'average' },
        { op: 'joinSheets', source: '订单!A1:B4', sourceKey: 'A', lookup: '客户!A1:C3', lookupKey: 'A', valueColumns: 'B', outputColumns: 'C' },
      ],
    }],
  }))
  const plan = await planner.plan({
    goal: '交叉表',
    path: 'D:/x.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '',
    validationSummary: '',
  })
  const steps = Array.isArray(plan) ? plan : plan.steps
  const crosstab = steps[0]!.operations[0] as Record<string, unknown>
  assert.deepEqual(crosstab.metric, { column: 'C', function: 'average' })
  const join = steps[0]!.operations[1] as Record<string, unknown>
  assert.deepEqual(join.valueColumns, ['B'])
  assert.deepEqual(join.outputColumns, ['C'])
})

test('planner passes machine assertions through when the model emits them', async () => {
  const planner = createLlmPlanner(async () => JSON.stringify({
    steps: [{ name: 's', operations: [{ op: 'aggregateReport', source: '订单!A1:C4', groupColumn: 'A', metrics: [{ column: 'C', function: 'sum' }] }] }],
    assertions: [{ id: '汇总!B2', startsWith: '=SUMIFS(' }],
  }))
  const plan = await planner.plan({
    goal: '按区域汇总金额',
    path: 'D:/x.xlsx',
    round: 1,
    sheetNames: ['订单'],
    profileSummary: '',
    validationSummary: '',
  })
  assert.ok(!Array.isArray(plan))
  assert.deepEqual(plan.assertions, [{ id: '汇总!B2', startsWith: '=SUMIFS(' }])
})

/**
 * The examples in the planner prompt have to be valid themselves.
 *
 * They are the shape the model copies, so an example that fails its own schema
 * teaches a failure: five of them wrote `"keep":"first|last"`, `"mode":"value|forward|left"`,
 * `"role":"ops|product|data"` and friends — placeholder syntax a model can paste
 * verbatim, and then the plan is rejected as an argument error. The prompt was
 * causing the class of failure it was meant to prevent.
 *
 * The check is mechanical: pull every JSON object out of the prompt file, and
 * require that each one already satisfies the schema it will be validated against.
 */
test('every example in the planner prompt satisfies the operation schema', async () => {
  const { readFile } = await import('node:fs/promises')
  const { fileURLToPath } = await import('node:url')
  const source = await readFile(fileURLToPath(new URL('../src/llm-planner.ts', import.meta.url)), 'utf8')

  // Pull out the first balanced object after each `'name: ` in the examples block.
  const examples = []
  for (const match of source.matchAll(/^\s*'([A-Za-z]+(?:\/[A-Za-z]+)*):\s*(\{)/gm)) {
    const start = match.index + match[0].length - 1
    let depth = 0
    let end = -1
    for (let i = start; i < source.length; i++) {
      if (source[i] === '{') depth++
      else if (source[i] === '}') { depth--; if (depth === 0) { end = i; break } }
    }
    if (end < 0) continue
    try {
      examples.push({ label: match[1], value: JSON.parse(source.slice(start, end + 1)) })
    } catch {
      // Not a JSON literal (a snippet with prose inside); the schema check below
      // only applies to the ones that are.
    }
  }
  assert.ok(examples.length >= 15, 'expected the prompt to carry worked examples, found ' + examples.length)

  const problems = []
  const branches = excelOperationSchema.oneOf
  for (const example of examples) {
    const op = example.value.op
    const branch = branches.find((b) => b.properties.op.enum[0] === op)
    if (!branch) { problems.push(example.label + ': op ' + JSON.stringify(op) + ' is not in the schema'); continue }
    for (const [key, spec] of Object.entries(branch.properties)) {
      if (key === 'op') continue
      const present = example.value[key] !== undefined
      if (spec.required === true && !present) { problems.push(example.label + ': missing required field ' + key); continue }
      if (!present) continue
      if (Array.isArray(spec.enum) && !spec.enum.includes(example.value[key])) {
        problems.push(example.label + ': ' + key + '=' + JSON.stringify(example.value[key]) + ' is not one of ' + spec.enum.join('/'))
      }
      if (spec.type === 'boolean' && typeof example.value[key] !== 'boolean') {
        problems.push(example.label + ': ' + key + '=' + JSON.stringify(example.value[key]) + ' should be a boolean')
      }
    }
  }
  assert.deepEqual(problems, [])
})
