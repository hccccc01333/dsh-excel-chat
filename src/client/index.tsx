/**
 * dsh-excel-chat browser half: registers per-tool renderers into the keyed
 * `tool.call.toolview` seat, so selecting an excel_* tool row in chat shows a
 * real spreadsheet grid (x-data-spreadsheet) in the right details column
 * instead of raw JSON or a text table.
 *
 * Editing: the grid is live-editable; committing a cell fills the composer
 * draft through `inputActions.setDraft` and submits it, so the agent applies
 * the change via excel_operate.set and re-previews. excel_autofix / excel_task
 * results render as an old→new repair diff.
 */
import { useEffect, useRef, useState, type ReactNode } from 'react'
import * as XSpreadsheetModule from 'x-data-spreadsheet/dist/xspreadsheet.js'
import xspreadsheetCss from 'x-data-spreadsheet/dist/xspreadsheet.css'
import type { ClientContext } from '@deepseek-ai/dsh-client-runtime/client'
import type { ToolCallViewProps } from '@deepseek-ai/dsh-client-ui-tool/client'
import type {} from '@deepseek-ai/dsh-client-ui-slots'
import type {} from '@deepseek-ai/dsh-client-ui-conversation/client'

/** Tools whose settled result is a cell table (excel_read / excel_preview). */
const TABLE_TOOLS = new Set(['excel_read', 'excel_preview'])
/** Tools whose settled result is a report/summary object. */
const SUMMARY_TOOLS = new Set(['excel_insight', 'excel_operate', 'excel_autofix', 'excel_task', 'excel_menu'])
/** Tools whose result carries repair diffs (old → new). */
const DIFF_TOOLS = new Set(['excel_autofix', 'excel_task'])

export const inject = ['slots']

interface SpreadsheetInstance {
  loadData(data: unknown): void
  on(event: string, callback: (...args: unknown[]) => void): void
  getCellText(row: number, col: number): string
  destroy(): void
}

// UMD build exposes a factory (window.x_spreadsheet / module default): call it
// WITHOUT `new`, it returns the sheet instance.
const Spreadsheet = (
  (window as unknown as { x_spreadsheet?: unknown }).x_spreadsheet
  ?? (XSpreadsheetModule as { default?: unknown }).default
  ?? XSpreadsheetModule
) as (el: HTMLElement, options: Record<string, unknown>) => SpreadsheetInstance

if (typeof document !== 'undefined' && document.querySelector('style[data-plugin="dsh-excel-chat"]') === null) {
  const style = document.createElement('style')
  style.dataset.plugin = 'dsh-excel-chat'
  style.textContent = xspreadsheetCss
  document.head.appendChild(style)
}

interface TableShape {
  sheets?: Array<{ sheet: string; cells: Array<{ id: string; value: unknown; formula?: string; type?: string }> }>
  markdown?: string
}

interface RepairLike {
  id?: unknown
  kind?: unknown
  oldValue?: unknown
  newValue?: unknown
}

function settledText(block: ToolCallViewProps['block']): string | null {
  if (!('kind' in block)) return null
  const parts = block.content.map((item) => (item.type === 'text' ? item.text : JSON.stringify(item)))
  return parts.join('\n') || null
}

function cellText(value: unknown, formula?: string): string {
  if (formula) return formula
  if (value === null || value === undefined) return ''
  if (typeof value === 'object') return JSON.stringify(value)
  return String(value)
}

/** Real spreadsheet grid (Excel-like) for excel_read / excel_preview results. */
function SpreadsheetView({
  sheets,
  editable,
  onCellEdited,
}: {
  sheets: NonNullable<TableShape['sheets']>
  editable: boolean
  onCellEdited: (cellId: string, text: string) => void
}): ReactNode {
  const ref = useRef<HTMLDivElement | null>(null)
  const ssRef = useRef<{ destroy: () => void } | null>(null)
  const [initError, setInitError] = useState<string | null>(null)
  useEffect(() => {
    const el = ref.current
    if (!el) return
    try {
      const maxRows = Math.max(...sheets.map((sheet) => {
        const rows = sheet.cells.map((cell) => Number(/(\d+)$/.exec(cell.id.split('!').pop() ?? '')?.[1] ?? 0))
        return rows.length > 0 ? Math.max(...rows) : 1
      }))
      const maxCols = Math.max(...sheets.map((sheet) => {
        const cols = sheet.cells.map((cell) => {
          const match = /([A-Za-z]+)\d+$/.exec(cell.id.split('!').pop() ?? '')
          return match
            ? [...match[1]!].reduce((acc, char) => acc * 26 + (char.toUpperCase().charCodeAt(0) - 64), 0)
            : 0
        })
        return cols.length > 0 ? Math.max(...cols) : 1
      }))
      const colCount = Math.max(5, maxCols)
      const rowCount = Math.min(200, Math.max(10, maxRows))
      const ss = Spreadsheet(el, {
        mode: editable ? 'edit' : 'read',
        showToolbar: false,
        showGrid: true,
        row: { len: rowCount, height: 24 },
        col: { len: colCount, width: 110 },
        view: { height: () => 420 },
      })
      const data = sheets.map((sheet) => {
        const rows: Record<number, { cells: Record<number, { text: string }> }> = {}
        for (const cell of sheet.cells) {
          const body = cell.id.split('!').pop() ?? ''
          const colMatch = /([A-Za-z]+)\d+$/.exec(body)
          const rowMatch = /(\d+)$/.exec(body)
          if (!colMatch || !rowMatch) continue
          const colIndex = [...colMatch[1]!].reduce((acc, char) => acc * 26 + (char.toUpperCase().charCodeAt(0) - 64), 0) - 1
          const rowIndex = Number(rowMatch[1]!) - 1
          const rowsOf = (rows[rowIndex] ??= { cells: {} })
          rowsOf.cells[colIndex] = { text: cellText(cell.value, cell.formula) }
        }
        return { name: sheet.sheet, rows }
      })
      ss.loadData(data)
      if (editable) {
        ss.on('cell-edited', (_cell, rowIndex: number, colIndex: number) => {
          const sheetName = sheets[0]?.sheet ?? 'Sheet1'
          const column = String.fromCharCode(65 + colIndex)
          const text = ss.getCellText(rowIndex, colIndex)
          onCellEdited(`${sheetName}!${column}${rowIndex + 1}`, text)
        })
      }
      ssRef.current = ss
      setInitError(null)
    } catch (error) {
      setInitError(error instanceof Error ? error.message : String(error))
    }
    return () => {
      ssRef.current = null
      el.innerHTML = ''
    }
  }, [sheets, editable])
  return (
    <div>
      <div ref={ref} style={{ border: '1px solid #d0d7de', borderRadius: 6, overflow: 'hidden' }} />
      {initError !== null && (
        <div style={{ fontSize: 11, color: '#b91c1c', marginTop: 6 }}>网格初始化失败：{initError}</div>
      )}
    </div>
  )
}

/** M3: render repair diffs (old → new) from excel_autofix / excel_task results. */
function DiffFromResult(result: unknown): ReactNode {
  const record = result as Record<string, unknown>
  const repairs = Array.isArray(record.repairs) ? (record.repairs as RepairLike[]) : []
  const rows: Array<{ id: string; old: string; next: string }> = []
  for (const repair of repairs) {
    if (typeof repair.id !== 'string') continue
    rows.push({
      id: repair.id,
      old: String(repair.oldValue ?? ''),
      next: String(repair.newValue ?? ''),
    })
  }
  const stepRows: string[] = []
  if (Array.isArray(record.steps)) {
    for (const step of record.steps as Array<{ name?: unknown; validation?: { before?: unknown; after?: unknown; fixed?: unknown } }>) {
      const validation = step.validation
      if (validation) {
        stepRows.push(`${String(step.name ?? '')}: 修复前 ${String(validation.before ?? '?')} 异常 → 修复后 ${String(validation.after ?? '?')}（修复 ${String(validation.fixed ?? 0)} 处）`)
      }
    }
  }
  const lines: string[] = []
  if (typeof record.message === 'string') lines.push(record.message)
  if (typeof record.repairedPath === 'string') lines.push(`输出：${record.repairedPath}`)
  if (typeof record.outputPath === 'string') lines.push(`输出：${record.outputPath}`)
  const nodes: ReactNode[] = []
  if (lines.length > 0) {
    nodes.push(<pre key="msg" style={{ fontSize: 12, whiteSpace: 'pre-wrap', margin: '0 0 8px' }}>{lines.join('\n')}</pre>)
  }
  if (stepRows.length > 0) {
    nodes.push(<pre key="steps" style={{ fontSize: 12, whiteSpace: 'pre-wrap', margin: '0 0 8px' }}>{stepRows.join('\n')}</pre>)
  }
  if (rows.length > 0) {
    const cellStyle = (header: boolean): React.CSSProperties => ({
      border: '1px solid #ddd',
      padding: '4px 8px',
      textAlign: 'left',
      fontSize: 12,
      background: header ? '#f3f4f6' : undefined,
    })
    nodes.push(
      <table key="diff" style={{ borderCollapse: 'collapse', width: '100%' }}>
        <thead>
          <tr>
            {['单元格', '修复前', '修复后'].map((label) => <th key={label} style={cellStyle(true)}>{label}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.id}>
              <td style={cellStyle(false)}>{row.id}</td>
              <td style={{ ...cellStyle(false), textDecoration: 'line-through', color: '#b91c1c' }}>{row.old}</td>
              <td style={{ ...cellStyle(false), color: '#15803d' }}>{row.next}</td>
            </tr>
          ))}
        </tbody>
      </table>,
    )
  }
  return nodes.length > 0 ? <div>{nodes}</div> : null
}

/** Toolview for one excel_* tool: spreadsheet grid, repair diff, or summary. */
export function ExcelToolView(props: ToolCallViewProps): ReactNode {
  const { toolName, block, inputActions } = props
  const text = settledText(block)
  if (text === null) return null
  let parsed: unknown = null
  try {
    parsed = JSON.parse(text)
  } catch {
    // Plain-text result (e.g. running fragment) — render as-is.
  }
  const canEdit = TABLE_TOOLS.has(toolName) && typeof inputActions?.setDraft === 'function'
  if (TABLE_TOOLS.has(toolName) && parsed !== null) {
    const value = parsed as TableShape
    if (Array.isArray(value.sheets) && value.sheets.length > 0) {
      return (
        <div>
          <SpreadsheetView
            sheets={value.sheets}
            editable={canEdit}
            onCellEdited={(cellId, newValue) => {
              if (!newValue) return
              const instruction = `请把 ${cellId} 的值改成 ${newValue}，用 excel_operate 的 set 执行，完成后用 excel_read 重新预览并回复“完成”。`
              inputActions.setDraft(instruction)
              inputActions.submit()
            }}
          />
          {canEdit && (
            <div style={{ fontSize: 11, color: '#6b7280', marginTop: 6 }}>
              改完直接回车：变更会通过对话让 agent 执行 set 并重新预览
            </div>
          )}
        </div>
      )
    }
    if (typeof value.markdown === 'string') {
      return <pre style={{ fontSize: 12, overflow: 'auto', whiteSpace: 'pre-wrap', maxHeight: 480 }}>{value.markdown}</pre>
    }
  }
  if (DIFF_TOOLS.has(toolName) && parsed !== null && typeof parsed === 'object') {
    const diff = DiffFromResult(parsed)
    if (diff !== null) return diff
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
      return <pre style={{ fontSize: 12, overflow: 'auto', whiteSpace: 'pre-wrap', maxHeight: 480 }}>{[...lines, ...counts].join('\n')}</pre>
    }
  }
  return <pre style={{ fontSize: 12, overflow: 'auto', whiteSpace: 'pre-wrap', maxHeight: 480 }}>{text}</pre>
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
