import { test } from 'node:test'
import assert from 'node:assert/strict'
import { ENUM_ALIASES, buildHeaderIndex, sanitizeAssertions, sanitizeOperations, sanitizePlan } from '../src/plan-schema.ts'
import { excelOperationSchema } from '../src/operation-schema.ts'

test('sanitizeOperations gives a tool call the same error the planner gets', () => {
  // `excel_operate` hands the model's array straight to the executor, which
  // assumes it is well-formed. Without this pass a missing nested field reached
  // a handler and surfaced as `TypeError: Cannot read properties of undefined
  // (reading 'toUpperCase')` — a message that names neither the operation nor
  // the field. The planner path already produced the actionable version.
  assert.throws(
    () => sanitizeOperations(
      [{ op: 'sortRange', range: 'Sheet1!A1:B3', keys: [{ direction: 'desc' }] }] as never,
      ['Sheet1'],
    ),
    /keys\[0\]\.column 缺失/,
  )
})

test('sanitizeOperations salvages exactly like sanitizePlan', () => {
  const operations = sanitizeOperations(
    [{ op: 'style', range: 'A1:B1', style: { bold: true } }] as never,
    ['订单'],
  )
  assert.equal((operations[0] as { range: string }).range, '订单!A1:B1')
})

test('sanitizePlan prefixes bare ranges with the first sheet', () => {
  const { steps } = sanitizePlan([{
    name: 's',
    operations: [{ op: 'style', range: 'A1:B1', style: { bold: true } }],
  }], ['订单'])
  assert.equal((steps[0]!.operations[0] as { range: string }).range, '订单!A1:B1')
})

test('sanitizePlan fills missing sheet and wraps single objects into arrays', () => {
  const { steps, notes } = sanitizePlan([{
    name: 's',
    operations: [
      { op: 'dedupeRows', columns: ['A'] },
      { op: 'sortRange', range: '订单!A1:B4', keys: { column: 'B', direction: 'desc' } },
    ],
  }], ['订单'])
  assert.equal((steps[0]!.operations[0] as { sheet: string }).sheet, '订单')
  assert.ok(Array.isArray((steps[0]!.operations[1] as { keys: unknown }).keys))
  assert.ok(notes.length >= 1)
})

test('sanitizePlan applies fillMissing aliases', () => {
  const { steps } = sanitizePlan([{
    name: 's',
    operations: [{ op: 'fillMissing', range: '订单!B2:B4', fillValue: 0 }],
  }], ['订单'])
  const op = steps[0]!.operations[0] as { mode: string; value: number }
  assert.equal(op.mode, 'value')
  assert.equal(op.value, 0)
})

test('sanitizePlan throws when a required array field is missing', () => {
  assert.throws(() => sanitizePlan([{
    name: 's',
    operations: [{ op: 'aggregateReport', source: '订单!A1:B4', groupColumn: 'A' }],
  }], ['订单']), /缺少必填数组 metrics/)
})

test('sanitizePlan requires freezePanes row and column', () => {
  assert.throws(
    () => sanitizePlan([{ operations: [{ op: 'freezePanes', sheet: '订单', range: '订单!A1:B2' }] }], ['订单']),
    /freezePanes 缺少必填数字 row/,
  )
  assert.throws(
    () => sanitizePlan([{ operations: [{ op: 'freezePanes', sheet: '订单', row: 2 }] }], ['订单']),
    /freezePanes 缺少必填字段 column/,
  )
  const fixed = sanitizePlan([{ operations: [{ op: 'freezePanes', sheet: '订单', column: 'A', row: '2' }] }], ['订单'])
  assert.equal((fixed.steps[0]!.operations[0] as Record<string, unknown>).row, 2)
})

test('sanitizePlan expands a single-cell fillSeries target to a range', () => {
  const fixed = sanitizePlan(
    [{ operations: [{ op: 'fillSeries', start: '订单!A2', target: '订单!A3' }] }],
    ['订单'],
  )
  assert.equal((fixed.steps[0]!.operations[0] as Record<string, unknown>).target, '订单!A2:A3')
})

test('sanitizePlan derives freezePanes row and column from a single-cell range', () => {
  const fixed = sanitizePlan(
    [{ operations: [{ op: 'freezePanes', sheet: '订单', range: '订单!A2' }] }],
    ['订单'],
  )
  const operation = fixed.steps[0]!.operations[0] as Record<string, unknown>
  assert.equal(operation.column, 'A')
  assert.equal(operation.row, 2)
})

test('sanitizePlan validates joinSheets arrays and key fields', () => {
  assert.throws(
    () => sanitizePlan([{
      operations: [{ op: 'joinSheets', source: '订单!A1:B4', sourceKey: 'A', lookup: '客户!A1:C3', lookupKey: 'A', valueColumns: ['B'] }],
    }], ['订单']),
    /缺少必填数组 outputColumns/,
  )
  assert.throws(
    () => sanitizePlan([{
      operations: [{ op: 'joinSheets', source: '订单!A1:B4', sourceKey: 'A', lookup: '客户!A1:C3', valueColumns: ['B'], outputColumns: ['C'] }],
    }], ['订单']),
    /缺少必填字段 lookupKey/,
  )
  const fixed = sanitizePlan([{
    operations: [{ op: 'joinSheets', source: '订单!A1:B4', sourceKey: 'A', lookup: '客户!A1:C3', lookupKey: 'A', valueColumns: ['B'], outputColumns: ['C'] }],
  }], ['订单'])
  assert.ok(fixed.steps.length === 1)
})

test('sanitizePlan normalizes crosstab flat metric fields and validates function', () => {
  const fixed = sanitizePlan([{
    operations: [{ op: 'crosstab', source: '订单!A1:C4', rowColumn: 'A', columnColumn: 'B', metricColumn: 'C' }],
  }], ['订单'])
  const op = fixed.steps[0]!.operations[0] as Record<string, unknown>
  assert.deepEqual(op.metric, { column: 'C', function: 'sum' })
  assert.throws(
    () => sanitizePlan([{
      operations: [{ op: 'crosstab', source: '订单!A1:C4', rowColumn: 'A', columnColumn: 'B', metric: { column: 'C', function: 'median' } }],
    }], ['订单']),
    /metric.function 不支持/,
  )
})

test('sanitizePlan checks rankColumn and new layout ops', () => {
  assert.throws(
    () => sanitizePlan([{ operations: [{ op: 'rankColumn', range: '订单!A1:B4', metricColumn: 'B' }] }], ['订单']),
    /rankColumn 缺少必填字段 outputColumn/,
  )
  assert.throws(
    () => sanitizePlan([{ operations: [{ op: 'hideRows', sheet: '订单', from: 2 }] }], ['订单']),
    /hideRows 缺少必填数字 to/,
  )
  const fixed = sanitizePlan([{
    operations: [{ op: 'moveSheet', name: '汇总', position: '1' }],
  }], ['订单'])
  assert.equal((fixed.steps[0]!.operations[0] as Record<string, unknown>).position, 1)
})

test('sanitizePlan salvages colloquial sheet names onto the exact sheet list', () => {
  const fixed = sanitizePlan([{
    operations: [{ op: 'renameSheet', oldName: '订单表', newName: '销售表' }],
  }], ['订单'])
  const op = fixed.steps[0]!.operations[0] as Record<string, unknown>
  assert.equal(op.oldName, '订单')
  assert.equal(op.newName, '销售表')
  const deduped = sanitizePlan([{
    operations: [{ op: 'dedupeRows', sheet: '订单表', columns: ['A'] }],
  }], ['订单'])
  assert.equal((deduped.steps[0]!.operations[0] as Record<string, unknown>).sheet, '订单')
  // Unknown names pass through untouched.
  const untouched = sanitizePlan([{
    operations: [{ op: 'renameSheet', oldName: '不存在的表', newName: 'x' }],
  }], ['订单'])
  assert.equal((untouched.steps[0]!.operations[0] as Record<string, unknown>).oldName, '不存在的表')
})

test('sanitizePlan aliases exceljs-native alignment names in style', () => {
  const fixed = sanitizePlan([{
    operations: [{ op: 'style', range: '订单!A1:B1', style: { wrapText: true, horizontal: 'center', vertical: 'middle' } }],
  }], ['订单'])
  const style = (fixed.steps[0]!.operations[0] as Record<string, unknown>).style as Record<string, unknown>
  assert.equal(style.hAlign, 'center')
  assert.equal(style.vAlign, 'middle')
})

test('sanitizeAssertions keeps well-formed assertions and repairs salvageable ones', () => {
  const { assertions, notes } = sanitizeAssertions([
    { id: '汇总!B2', startsWith: '=SUMIFS(' },
    { id: 'B3', expect: 0 },
    { id: '订单!A1', expect: null },
    { expect: 1 },
    { id: '', expect: 'x' },
    { id: '订单!C1' },
    { id: '订单!D1', startsWith: '' },
    'not-an-object',
  ], ['订单'])
  assert.deepEqual(assertions, [
    { id: '汇总!B2', startsWith: '=SUMIFS(' },
    { id: '订单!B3', expect: '0' },
    { id: '订单!A1', expect: null },
  ])
  assert.ok(notes.length >= 4)
})

test('sanitizeAssertions ignores a non-array field entirely', () => {
  const { assertions, notes } = sanitizeAssertions('nope', ['订单'])
  assert.deepEqual(assertions, [])
  assert.ok(notes.some((note) => note.includes('不是数组')))
})

/**
 * Header names where column letters belong.
 *
 * A planner reading the profile sees `区域`, `金额` — the headers — so it writes
 * `groupColumn: "区域"` where the schema wants `"B"`. That failed as
 * `invalid column letter: 区域` even though the intent is unambiguous and the
 * answer is in the profile the agent already holds. Resolving it is a pure
 * salvage, so it belongs here rather than in a prompt the model may ignore.
 */
const HEADERS = buildHeaderIndex([
  { sheet: '订单', columns: [
    { column: 'A', header: '订单号' },
    { column: 'B', header: '区域' },
    { column: 'E', header: '数量' },
    // `金额` appears twice on purpose: the name no longer identifies one column,
    // so the resolver has to decline it rather than pick.
    { column: 'F', header: '金额' },
    { column: 'G', header: '  金额  ' },
  ] },
  { sheet: '价目表', columns: [
    { column: 'A', header: '产品' },
    { column: 'B', header: '单价' },
  ] },
])

test('a header name is resolved to the column letter the schema wants', () => {
  const { steps, notes } = sanitizePlan([{
    operations: [
      { op: 'subtotal', sheet: '订单', range: '订单!A1:F25', groupColumn: '区域', summaryColumns: [{ column: '数量', function: 'sum' }] },
    ],
  }] as never, ['订单', '价目表'], HEADERS)

  const operation = steps[0]!.operations[0] as Record<string, unknown>
  assert.equal(operation.groupColumn, 'B')
  assert.deepEqual((operation.summaryColumns as Array<{ column: string }>)[0]!.column, 'E')
  assert.ok(notes.some((note) => note.includes('区域')))
})

test('a column letter is left alone, and an unknown header is not invented', () => {
  const { steps } = sanitizePlan([{
    operations: [
      { op: 'subtotal', sheet: '订单', range: '订单!A1:F25', groupColumn: 'B', summaryColumns: [{ column: 'F', function: 'sum' }] },
      { op: 'subtotal', sheet: '订单', range: '订单!A1:F25', groupColumn: '不存在的列', summaryColumns: [{ column: 'F', function: 'sum' }] },
    ],
  }] as never, ['订单', '价目表'], HEADERS)

  const [letter, unknown] = steps[0]!.operations as Array<Record<string, unknown>>
  assert.equal(letter!.groupColumn, 'B', 'a valid letter must not be touched')
  // Left as-is so the failure stays exactly as loud as it was.
  assert.equal(unknown!.groupColumn, '不存在的列')
})

test('an ambiguous header is refused rather than guessed', () => {
  // Two columns share the header `金额`, so the name no longer identifies one
  // column. Picking either would write a plausible-looking wrong answer.
  const { steps } = sanitizePlan([{
    operations: [
      { op: 'subtotal', sheet: '订单', range: '订单!A1:F25', groupColumn: '金额', summaryColumns: [{ column: 'F', function: 'sum' }] },
    ],
  }] as never, ['订单'], HEADERS)

  assert.equal((steps[0]!.operations[0] as Record<string, unknown>).groupColumn, '金额')
})

test('a header resolves against the sheet the operation targets', () => {
  // `单价` is column B on 价目表 and absent from 订单, so the sheet has to be
  // honoured rather than looking everything up in the first sheet.
  const { steps } = sanitizePlan([{
    operations: [
      { op: 'sortRange', sheet: '价目表', range: '价目表!A1:B9', keys: [{ column: '单价', ascending: true }] },
    ],
  }] as never, ['订单', '价目表'], HEADERS)

  assert.equal((steps[0]!.operations[0] as Record<string, unknown> & { keys: Array<{ column: string }> }).keys[0]!.column, 'B')
})

test('header matching tolerates surrounding and repeated whitespace', () => {
  const { steps } = sanitizePlan([{
    operations: [
      { op: 'subtotal', sheet: '订单', range: '订单!A1:F25', groupColumn: ' 区域 ', summaryColumns: [{ column: 'F', function: 'sum' }] },
    ],
  }] as never, ['订单'], HEADERS)

  assert.equal((steps[0]!.operations[0] as Record<string, unknown>).groupColumn, 'B')
})

/**
 * Enum and boolean values as the model actually writes them.
 *
 * `function: "SUM"`, `function: "求和"`, `keep: "FIRST"`, `bold: "true"` are the
 * same constants as the schema's, spelled differently. Rejecting them is a
 * pure argument failure, and the schema already knows every legal value — so the
 * normalisation is derived from the schema rather than hand-listed, which is what
 * keeps it from drifting when a field gains an option.
 */
test('an enum value is normalised regardless of case', () => {
  const { steps, notes } = sanitizePlan([{
    operations: [
      { op: 'aggregateReport', source: '订单!A1:F9', groupColumn: 'B', metrics: [{ column: 'F', function: 'SUM' }] },
      { op: 'conditionalFormatting', range: '订单!A1:F9', rules: [{ type: 'DataBar' }] },
      { op: 'dedupeRows', sheet: '订单', keyColumns: ['A'], keep: 'FIRST' },
    ],
  }] as never, ['订单'])

  const [aggregate, formatting, dedupe] = steps[0]!.operations as Array<Record<string, never>>
  assert.equal((aggregate as { metrics: Array<{ function: string }> }).metrics[0]!.function, 'sum')
  assert.equal((formatting as { rules: Array<{ type: string }> }).rules[0]!.type, 'dataBar')
  assert.equal((dedupe as { keep: string }).keep, 'first')
  assert.equal(notes.length, 3)
})

test('an enum value written in Chinese is normalised to the constant', () => {
  const { steps } = sanitizePlan([{
    operations: [
      { op: 'aggregateReport', source: '订单!A1:F9', groupColumn: 'B', metrics: [{ column: 'F', function: '求和' }] },
      { op: 'sortRange', range: '订单!A1:F9', keys: [{ column: 'F', direction: '降序' }] },
      { op: 'style', range: '订单!A1', style: { hAlign: '居中', vAlign: '居中' } },
    ],
  }] as never, ['订单'])

  const [aggregate, sort, style] = steps[0]!.operations as Array<Record<string, never>>
  assert.equal((aggregate as { metrics: Array<{ function: string }> }).metrics[0]!.function, 'sum')
  assert.equal((sort as { keys: Array<{ direction: string }> }).keys[0]!.direction, 'desc')
  // The same word means different constants depending on the axis.
  assert.deepEqual((style as { style: Record<string, string> }).style, { hAlign: 'center', vAlign: 'middle' })
})

test('a boolean written as a string becomes a boolean', () => {
  const { steps } = sanitizePlan([{
    operations: [
      { op: 'report', source: '订单!A1:F9', groupColumn: 'B', metrics: [{ column: 'F', function: 'sum' }], autoFilter: 'false', freezeHeader: '是' },
      { op: 'style', range: '订单!A1', style: { bold: 'true', italic: '否' } },
    ],
  }] as never, ['订单'])

  const [report, style] = steps[0]!.operations as Array<Record<string, never>>
  assert.equal((report as { autoFilter: unknown }).autoFilter, false)
  assert.equal((report as { freezeHeader: unknown }).freezeHeader, true)
  assert.deepEqual((style as { style: Record<string, unknown> }).style, { bold: true, italic: false })
})

test('a legal value is left alone and an unknown one is not guessed', () => {
  const { steps, notes } = sanitizePlan([{
    operations: [
      { op: 'aggregateReport', source: '订单!A1:F9', groupColumn: 'B', metrics: [{ column: 'F', function: 'sum' }] },
      // `中位数` is not an allowed function, and no alias claims it is. Inventing
      // one would write a valid-looking wrong number into the workbook.
      { op: 'aggregateReport', source: '订单!A1:F9', groupColumn: 'B', metrics: [{ column: 'F', function: '中位数' }] },
    ],
  }] as never, ['订单'])

  const [legal, unknown] = steps[0]!.operations as Array<{ metrics: Array<{ function: string }> }>
  assert.equal(legal!.metrics[0]!.function, 'sum')
  assert.equal(unknown!.metrics[0]!.function, '中位数')
  assert.equal(notes.length, 0, 'nothing should have been rewritten')
})

/**
 * Every alias has to land on a value its field actually allows.
 *
 * The alias table is the one hand-written part of the normalisation, so it is the
 * one part that can rot: a typo'd target would make the alias silently do nothing,
 * and a target borrowed from the wrong field would rewrite a legal value into one
 * the schema rejects. Checked against the schema rather than against the table.
 */
test('every enum alias targets a value its field allows', () => {
  const allowedByOp = new Map<string, Map<string, Set<string>>>()
  for (const branch of excelOperationSchema.oneOf as unknown as Array<Record<string, unknown>>) {
    const properties = branch.properties as Record<string, Record<string, unknown>>
    const op = (properties.op.enum as string[])[0]!
    const fields = new Map<string, Set<string>>()
    const walk = (spec: unknown, field: string): void => {
      if (!spec || typeof spec !== 'object') return
      const node = spec as Record<string, unknown>
      if (Array.isArray(node.enum)) {
        if (!fields.has(field)) fields.set(field, new Set(node.enum as string[]))
      }
      for (const [key, child] of Object.entries(node)) {
        if (key === 'enum' || !child || typeof child !== 'object') continue
        walk(child, key === 'properties' || key === 'items' ? field : key)
      }
    }
    for (const [key, spec] of Object.entries(properties)) {
      if (key !== 'op') walk(spec, key)
    }
    allowedByOp.set(op, fields)
  }

  const problems: string[] = []
  for (const [field, aliases] of Object.entries(ENUM_ALIASES)) {
    const allowedAnywhere = new Set<string>()
    for (const fields of allowedByOp.values()) {
      const values = fields.get(field)
      if (values) for (const value of values) allowedAnywhere.add(value)
    }
    assert.ok(allowedAnywhere.size > 0, `alias table covers field ${field}, which is not an enum in the schema`)
    for (const [alias, target] of Object.entries(aliases)) {
      if (!allowedAnywhere.has(target)) problems.push(`${field}: 「${alias}」 -> 「${target}」, which is not an allowed value`)
    }
  }
  assert.deepEqual(problems, [])
})
