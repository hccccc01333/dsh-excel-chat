/**
 * dsh-excel-chat browser half: registers per-tool renderers into the keyed
 * `tool.call.toolview` seat, so selecting an excel_* tool row in chat shows a
 * rendered table / summary in the right details column instead of raw JSON.
 * Pure client rendering from the conversation snapshot — no host RPC needed.
 */
import { createElement, type ReactNode } from 'react'
import type { ClientContext } from '@deepseek-ai/dsh-client-runtime/client'
import type { ToolCallViewProps } from '@deepseek-ai/dsh-client-ui-tool/client'
import type {} from '@deepseek-ai/dsh-client-ui-slots'
import type {} from '@deepseek-ai/dsh-client-ui-conversation/client'

/** Tools whose settled result is a cell table (excel_read / excel_preview). */
const TABLE_TOOLS = new Set(['excel_read', 'excel_preview'])
/** Tools whose settled result is a report/summary object. */
const SUMMARY_TOOLS = new Set(['excel_insight', 'excel_operate', 'excel_autofix', 'excel_task', 'excel_menu'])

export const inject = ['slots']

interface TableShape {
  sheets?: Array<{ sheet: string; cells: Array<{ id: string; value: unknown; formula?: string; type?: string }> }>
  markdown?: string
}

function cellText(value: unknown, formula?: string): string {
  if (formula) return formula
  if (value === null || value === undefined) return ''
  if (typeof value === 'object') return JSON.stringify(value)
  return String(value)
}

/** Minimal table renderer for excel_read/excel_preview results. */
function TableFromResult(result: unknown): ReactNode {
  const value = result as TableShape
  const sheets = Array.isArray(value.sheets) ? value.sheets : null
  if (sheets) {
    const first = sheets[0]
    if (!first) return null
    const columns = new Set<string>()
    for (const cell of first.cells) {
      const match = /([A-Za-z]+)\d+$/.exec(cell.id.split('!').pop() ?? '')
      if (match) columns.add(match[1]!)
    }
    const columnList = [...columns].sort((a, b) => a.length - b.length || a.localeCompare(b))
    const byCell = new Map(first.cells.map((cell) => [cell.id, cell]))
    const rows = new Set<number>()
    for (const cell of first.cells) {
      const match = /(\d+)$/.exec(cell.id.split('!').pop() ?? '')
      if (match) rows.add(Number(match[1]!))
    }
    const rowList = [...rows].sort((a, b) => a - b)
    return createElement(
      'div',
      { style: { overflow: 'auto', maxHeight: 480 } },
      createElement('div', { style: { fontWeight: 600, marginBottom: 6 } }, first.sheet),
      createElement(
        'table',
        { style: { borderCollapse: 'collapse', fontSize: 12, width: '100%' } },
        createElement(
          'thead',
          null,
          createElement(
            'tr',
            null,
            columnList.map((column) =>
              createElement('th', { key: column, style: tableCellStyle(true) }, column)),
          ),
        ),
        createElement(
          'tbody',
          null,
          rowList.map((row) =>
            createElement(
              'tr',
              { key: row },
              columnList.map((column) => {
                const cell = byCell.get(`${first.sheet}!${column}${row}`)
                return createElement(
                  'td',
                  { key: `${column}${row}`, style: tableCellStyle(false) },
                  cell ? cellText(cell.value, cell.formula) : '',
                )
              }),
            )),
        ),
      ),
    )
  }
  if (typeof value.markdown === 'string') {
    return createElement(
      'pre',
      { style: { fontSize: 12, overflow: 'auto', whiteSpace: 'pre-wrap', maxHeight: 480 } },
      value.markdown,
    )
  }
  return null
}

function tableCellStyle(header: boolean): React.CSSProperties {
  return {
    border: '1px solid #ddd',
    padding: '4px 8px',
    textAlign: 'left',
    background: header ? '#f3f4f6' : undefined,
  }
}

function settledText(block: ToolCallViewProps['block']): string | null {
  if (!('kind' in block)) return null
  const parts = block.content.map((item) => (item.type === 'text' ? item.text : JSON.stringify(item)))
  return parts.join('\n') || null
}

/** Toolview for one excel_* tool: parse the settled result and render it. */
export function ExcelToolView({ toolName, block }: ToolCallViewProps): ReactNode {
  const text = settledText(block)
  if (text === null) return null
  let parsed: unknown = null
  try {
    parsed = JSON.parse(text)
  } catch {
    // Plain-text result (e.g. running fragment) — render as-is.
  }
  if (TABLE_TOOLS.has(toolName) && parsed !== null) {
    const table = TableFromResult(parsed)
    if (table !== null) return table
  }
  if (SUMMARY_TOOLS.has(toolName) && parsed !== null && typeof parsed === 'object') {
    const record = parsed as Record<string, unknown>
    const lines = ['message', 'summary', 'outputPath', 'repairedPath']
      .filter((key) => typeof record[key] === 'string')
      .map((key) => `${key}: ${String(record[key])}`)
    const counts = ['before', 'after']
      .filter((key) => record[key] && typeof record[key] === 'object')
      .map((key) => `${key}: ${JSON.stringify(record[key])}`)
    if (lines.length > 0 || counts.length > 0) {
      return createElement(
        'pre',
        { style: { fontSize: 12, overflow: 'auto', whiteSpace: 'pre-wrap', maxHeight: 480 } },
        [...lines, ...counts].join('\n'),
      )
    }
  }
  return createElement(
    'pre',
    { style: { fontSize: 12, overflow: 'auto', whiteSpace: 'pre-wrap', maxHeight: 480 } },
    text,
  )
}

/** Client plugin body: register the excel_* toolviews into the details column. */
export function apply(ctx: ClientContext): void {
  for (const tool of [...TABLE_TOOLS, ...SUMMARY_TOOLS]) {
    ctx.effect(
      () => ctx.slots.inject('tool.call.toolview', () => ctx.slots.register(
        { name: 'tool.call.toolview', key: tool },
        ExcelToolView,
      )),
      `dsh-excel-chat: toolview ${tool}`,
    )
  }
}
