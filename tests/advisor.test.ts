import { test } from 'node:test'
import assert from 'node:assert/strict'
import { createLlmRepairAdvisor, buildRepairPrompt, type LlmText } from '../src/advisor.ts'
import { applyPatches } from '../src/patch.ts'
import { generateRepairs } from '../src/repair.ts'
import { validate } from '../src/validator.ts'

const table = { sheet: 'Sheet1', columns: { revenue: 'B', cost: 'C' } }
const cells = {
  D2: '=B2-C2',
  D3: '=SUM(B3:C3)',
  D4: '=B4-C4',
}

test('buildRepairPrompt includes the excerpt, anomalies, and table schema', () => {
  const result = validate(cells)
  const prompt = buildRepairPrompt(cells, result.anomalies, table)
  assert.match(prompt, /D3/)
  assert.match(prompt, /structure-mismatch/)
  assert.match(prompt, /Sheet1/)
})

test('LLM advisor compiles IR repairs and the patched workbook re-validates clean', async () => {
  const result = validate(cells)
  assert.equal(generateRepairs(cells, result).length, 0)

  const fakeLlm: LlmText = async () => JSON.stringify({
    repairs: [{
      id: 'D3',
      baseCell: 'D3',
      ir: {
        operation: 'binary',
        left: { kind: 'column', column: 'revenue' },
        right: { kind: 'column', column: 'cost' },
        operator: '-',
      },
    }],
  })

  const advisor = createLlmRepairAdvisor(fakeLlm, table)
  const patches = await advisor(cells, result)
  assert.deepEqual(patches, [{
    id: 'D3',
    kind: 'formula',
    oldValue: '=SUM(B3:C3)',
    newValue: '=B3-C3',
  }])

  const fixed = applyPatches(cells, patches)
  const after = validate(fixed)
  assert.equal(after.anomalies.length, 0)
})

test('LLM advisor strips markdown code fences', async () => {
  const result = validate(cells)
  const fakeLlm: LlmText = async () => '```json\n' + JSON.stringify({
    repairs: [{
      id: 'D3',
      baseCell: 'D3',
      ir: {
        operation: 'binary',
        left: { kind: 'column', column: 'revenue' },
        right: { kind: 'column', column: 'cost' },
        operator: '-',
      },
    }],
  }) + '\n```'
  const advisor = createLlmRepairAdvisor(fakeLlm, table)
  const patches = await advisor(cells, result)
  assert.equal(patches.length, 1)
})

test('LLM advisor resolves sheet-qualified cell ids from bare ids', async () => {
  const sheetCells = {
    'Sheet1!D2': '=B2-C2',
    'Sheet1!D3': '=SUM(B3:C3)',
    'Sheet1!D4': '=B4-C4',
  }
  const result = validate(sheetCells)
  const fakeLlm: LlmText = async () => JSON.stringify({
    repairs: [{
      id: 'D3',
      baseCell: 'D3',
      ir: {
        operation: 'binary',
        left: { kind: 'column', column: 'revenue' },
        right: { kind: 'column', column: 'cost' },
        operator: '-',
      },
    }],
  })
  const advisor = createLlmRepairAdvisor(fakeLlm, table)
  const patches = await advisor(sheetCells, result)
  assert.equal(patches.length, 1)
  assert.equal(patches[0]!.id, 'Sheet1!D3')
  assert.equal(patches[0]!.newValue, '=B3-C3')
})
