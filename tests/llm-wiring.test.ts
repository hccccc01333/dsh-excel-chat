import { test } from 'node:test'
import assert from 'node:assert/strict'
import { Context } from '@deepseek-ai/cordis'
import { CallId, LlmAdapter, LlmRuntime, type GenerateOptions, type StreamChunk } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import ExcelJS from 'exceljs'
import { fileURLToPath } from 'node:url'
import * as plugin from '../src/index.ts'

const fixturePath = fileURLToPath(new URL('../fixtures/structure-mismatch.xlsx', import.meta.url))

class FakeLlmAdapter extends LlmAdapter {
  stream(_options: GenerateOptions): AsyncIterable<StreamChunk> {
    const reply = JSON.stringify({
      repairs: [{
        id: 'Sheet1!D3',
        baseCell: 'Sheet1!D3',
        ir: {
          operation: 'binary',
          left: { kind: 'column', column: 'revenue' },
          right: { kind: 'column', column: 'cost' },
          operator: '-',
        },
      }],
    })
    return (async function* () {
      yield { type: 'block-start', index: 0, blockType: 'text' }
      yield { type: 'text-delta', index: 0, text: reply }
      yield { type: 'block-end', index: 0, block: { type: 'text', text: reply } }
      yield { type: 'finish', reason: { kind: 'stop' } }
    })()
  }
}

test('excel_repair_formulas uses ctx.llm to repair structure mismatches', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('D2').value = { formula: 'B2-C2', result: 40 }
  sheet.getCell('D3').value = { formula: 'SUM(B3:C3)', result: 80 }
  sheet.getCell('D4').value = { formula: 'B4-C4', result: 80 }
  await workbook.xlsx.writeFile(fixturePath)

  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(LlmRuntime)
  ctx.llm.registerAdapter(['fake'], new FakeLlmAdapter())
  await ctx.plugin(plugin)

  const result = await ctx.tools.execute({
    signal: new AbortController().signal,
    callId: CallId('vera-llm-wiring-1'),
    name: 'excel_repair_formulas',
    arguments: {
      path: fixturePath,
      useLlm: true,
      provider: 'fake',
      model: 'fake-model',
      table: { sheet: 'Sheet1', columns: { revenue: 'B', cost: 'C' } },
    },
  })

  assert.equal(result.isError, false)
  const value = result.value as {
    llmRepairs: Array<{ id: string; oldValue: string; newValue: string }>
    before: { anomalies: Array<{ kind: string }> }
    after: { anomalies: Array<{ kind: string }> }
    repairedPath: string
  }
  assert.equal(value.llmRepairs.length, 1)
  assert.equal(value.llmRepairs[0]!.newValue, '=B3-C3')
  assert.ok(value.before.anomalies.some((a) => a.kind === 'structure-mismatch'))
  assert.equal(value.after.anomalies.filter((a) => a.kind === 'structure-mismatch').length, 0)
})
