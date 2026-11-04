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
import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import * as XSpreadsheetModule from 'x-data-spreadsheet/dist/xspreadsheet.js'
import xspreadsheetCss from 'x-data-spreadsheet/dist/xspreadsheet.css'
import type { ClientContext } from '@deepseek-ai/dsh-client-runtime/client'
import type { ToolCallViewProps } from '@deepseek-ai/dsh-client-ui-tool/client'
import type {} from '@deepseek-ai/dsh-client-ui-slots'
import type {} from '@deepseek-ai/dsh-client-ui-conversation/client'
import {
  FULLSCREEN_RENDER_COLS,
  FULLSCREEN_RENDER_ROWS,
  columnName,
  measureSheets,
  toGridData,
  type TableShape,
} from './view-model.ts'
import {
  captureWindowListeners,
  keepEditorFocusInPlace,
  scopeGlobalTypingToGrid,
  type WindowListener,
} from './dom-guards.ts'

/** Tools whose settled result is a cell table (excel_read / excel_preview). */
const TABLE_TOOLS = new Set(['excel_read', 'excel_preview'])
/** Tools whose settled result is a report/summary object. */
const SUMMARY_TOOLS = new Set(['excel_insight', 'excel_operate', 'excel_autofix', 'excel_task', 'excel_menu'])
/** Tools whose result carries repair diffs (old → new). */
const DIFF_TOOLS = new Set(['excel_autofix', 'excel_task'])

export const inject = ['slots', 'remote.commands']

interface GridInstance {
  loadData(data: unknown): void
  on(event: string, callback: (...args: unknown[]) => void): void
  /** Not on the public typings, but the library exposes its sheet object. */
  sheet?: { reload(): void }
}

// UMD build exposes a factory (window.x_spreadsheet / module default): call it
// WITHOUT `new`, it returns the sheet instance.
const Spreadsheet = (
  (window as unknown as { x_spreadsheet?: unknown }).x_spreadsheet
  ?? (XSpreadsheetModule as { default?: unknown }).default
  ?? XSpreadsheetModule
) as (el: HTMLElement, options: Record<string, unknown>) => GridInstance

interface CommandsRemote {
  execute(sessionId: string, line: string): Promise<{ result: { kind: string; text?: string } }>
}

/** Set once by the client plugin body; used by the toolview for in-place edits. */
let commandsRemote: CommandsRemote | null = null

if (typeof document !== 'undefined' && document.querySelector('style[data-plugin="dsh-excel-chat"]') === null) {
  const style = document.createElement('style')
  style.dataset.plugin = 'dsh-excel-chat'
  style.textContent = xspreadsheetCss
  document.head.appendChild(style)
}

interface RepairLike {
  id?: unknown
  kind?: unknown
  oldValue?: unknown
  newValue?: unknown
}

interface MenuSuggestionLike {
  id?: unknown
  title?: unknown
  description?: unknown
  example?: unknown
}

function settledText(block: ToolCallViewProps['block']): string | null {
  if (!('kind' in block)) return null
  const parts = block.content.map((item) => (item.type === 'text' ? item.text : JSON.stringify(item)))
  return parts.join('\n') || null
}

/** Inline viewport height; the fullscreen overlay uses most of the window. */
const INLINE_VIEW_HEIGHT = 560

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
  /** Current inner width of the column; the grid's view.width reads it. */
  const widthRef = useRef(0)
  const [initError, setInitError] = useState<string | null>(null)
  const [fullscreen, setFullscreen] = useState(false)
  const [viewportTick, setViewportTick] = useState(0)
  // Derived, not stored: a state copy would let the notice and the grid drift.
  const layout = useMemo(() => measureSheets(sheets, fullscreen), [sheets, fullscreen])
  const viewHeight = fullscreen ? Math.max(320, window.innerHeight - 96) : INLINE_VIEW_HEIGHT
  useEffect(() => {
    if (!fullscreen) return
    const onResize = (): void => setViewportTick((tick) => tick + 1)
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setFullscreen(false)
    }
    window.addEventListener('resize', onResize)
    window.addEventListener('keydown', onKeyDown)
    return () => {
      window.removeEventListener('resize', onResize)
      window.removeEventListener('keydown', onKeyDown)
    }
  }, [fullscreen])
  useEffect(() => {
    const el = ref.current
    if (el === null) return
    // Who had focus before the grid existed: mounting a card can pull focus into
    // it (the host focuses what it just mounted, the library focuses its
    // editor), and a canvas is nowhere to leave a caret.
    const focusedBefore = document.activeElement instanceof HTMLElement ? document.activeElement : null
    let listeners: WindowListener[] = []
    let observer: ResizeObserver | null = null
    try {
      widthRef.current = el.clientWidth || document.documentElement.clientWidth
      const built = captureWindowListeners(() => Spreadsheet(el, {
        mode: editable ? 'edit' : 'read',
        showToolbar: false,
        showGrid: true,
        row: { len: layout.rows, height: 24 },
        col: { len: layout.cols, width: 110 },
        // Width matters as much as height: the library's default is the whole
        // window, which draws a grid several times wider than the details
        // column and puts its horizontal scrollbar out of reach.
        view: { height: () => viewHeight, width: () => widthRef.current },
      }))
      const ss = built.value
      listeners = built.listeners
      ss.loadData(toGridData(sheets))
      if (editable) {
        ss.on('cell-edited', (editedCell, rowIndex, colIndex) => {
          const sheetName = sheets[0]?.sheet ?? 'Sheet1'
          const text = editedCell && typeof editedCell === 'object'
            ? String((editedCell as { text?: unknown }).text ?? '')
            : ''
          onCellEdited(`${sheetName}!${columnName(Number(colIndex) + 1)}${Number(rowIndex) + 1}`, text)
        })
      }
      scopeGlobalTypingToGrid(listeners, el)
      keepEditorFocusInPlace(el)
      // The column can be resized without the window resizing, and the grid
      // sizes its canvas from the width it saw at construction.
      observer = new ResizeObserver(() => {
        const width = el.clientWidth
        if (width <= 0 || width === widthRef.current) return
        widthRef.current = width
        ss.sheet?.reload()
      })
      observer.observe(el)
      setInitError(null)
      requestAnimationFrame(() => {
        if (!el.contains(document.activeElement)) return
        if (focusedBefore !== null && focusedBefore.isConnected) focusedBefore.focus({ preventScroll: true })
        else (document.activeElement as HTMLElement | null)?.blur()
      })
    } catch (error) {
      setInitError(error instanceof Error ? error.message : String(error))
    }
    return () => {
      if (observer !== null) observer.disconnect()
      for (const [type, listener, options] of listeners) window.removeEventListener(type, listener, options)
      el.innerHTML = ''
    }
  }, [sheets, editable, fullscreen, viewportTick, layout.rows, layout.cols, viewHeight])
  const truncationNotice = layout.truncated
    ? fullscreen
      ? `表格较大：已显示前 ${layout.capRows} 行 × ${layout.capCols} 列，已达全屏显示上限`
      : `表格较大：已显示前 ${layout.capRows} 行 × ${layout.capCols} 列，点「全屏」可看更多（${FULLSCREEN_RENDER_ROWS} 行 × ${FULLSCREEN_RENDER_COLS} 列）`
    : null
  return (
    <div
      style={fullscreen
        ? {
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            zIndex: 9999,
            background: '#fff',
            padding: 12,
            overflow: 'hidden',
          }
        : undefined}
    >
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 4 }}>
        <button
          type="button"
          onClick={() => setFullscreen((value) => !value)}
          style={{ fontSize: 11, cursor: 'pointer' }}
        >
          {fullscreen ? '退出全屏（Esc）' : '全屏'}
        </button>
      </div>
      <div
        ref={ref}
        style={fullscreen
          ? { overflow: 'hidden' }
          : {
              border: '1px solid #d0d7de',
              borderRadius: 6,
              overflow: 'hidden',
              // Reserve the grid's footprint: if the canvas is ever rebuilt, the
              // column must not collapse for a frame and drag the scroll
              // position of the conversation with it.
              minHeight: INLINE_VIEW_HEIGHT,
            }}
      />
      {truncationNotice !== null && (
        <div style={{ fontSize: 11, color: '#92400e', marginTop: 6 }}>{truncationNotice}</div>
      )}
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

/** Rendered capability menu for excel_menu results. */
function MenuFromResult(result: unknown): ReactNode {
  const record = result as Record<string, unknown>
  const suggestions = Array.isArray(record.suggestions) ? (record.suggestions as MenuSuggestionLike[]) : []
  const items = suggestions
    .map((suggestion, index) => ({
      id: String(suggestion.id ?? index + 1),
      title: String(suggestion.title ?? ''),
      description: String(suggestion.description ?? ''),
      example: String(suggestion.example ?? ''),
    }))
    .filter((item) => item.title !== '')
  const summary = typeof record.summary === 'string' ? record.summary : ''
  const nodes: ReactNode[] = []
  if (summary) nodes.push(<pre key="summary" style={{ fontSize: 12, whiteSpace: 'pre-wrap', margin: '0 0 8px' }}>{summary}</pre>)
  nodes.push(
    <ol key="menu" style={{ margin: 0, paddingLeft: 18, fontSize: 12 }}>
      {items.map((item) => (
        <li key={item.id} style={{ marginBottom: 6 }}>
          <div style={{ fontWeight: 600 }}>{item.title}</div>
          <div style={{ color: '#6b7280' }}>{item.description}</div>
          {item.example && <div style={{ color: '#2563eb', fontFamily: 'monospace' }}>例：{item.example}</div>}
        </li>
      ))}
    </ol>,
  )
  return <div>{nodes}</div>
}

/** Rendered findings list for excel_insight results. */
function InsightFromResult(result: unknown): ReactNode {
  const record = result as Record<string, unknown>
  const sheets = Array.isArray(record.sheets) ? record.sheets as Array<Record<string, unknown>> : []
  const findings: Array<{ severity: string; message: string }> = []
  for (const sheet of sheets) {
    const list = Array.isArray(sheet.findings) ? sheet.findings as Array<Record<string, unknown>> : []
    for (const finding of list) {
      findings.push({
        severity: String(finding.severity ?? 'info'),
        message: String(finding.message ?? ''),
      })
    }
  }
  const suggestions = Array.isArray(record.suggestions) ? (record.suggestions as unknown[]) : []
  const nodes: ReactNode[] = []
  if (typeof record.summary === 'string') {
    nodes.push(<pre key="summary" style={{ fontSize: 12, whiteSpace: 'pre-wrap', margin: '0 0 8px' }}>{record.summary}</pre>)
  }
  for (const finding of findings) {
    const color = finding.severity === 'alert' ? '#b91c1c' : finding.severity === 'warn' ? '#b45309' : '#1f2937'
    nodes.push(<div key={finding.message} style={{ fontSize: 12, color, marginBottom: 4 }}>• {finding.message}</div>)
  }
  if (suggestions.length > 0) {
    nodes.push(<pre key="suggestions" style={{ fontSize: 12, whiteSpace: 'pre-wrap', margin: '8px 0 0' }}>{suggestions.join('\n')}</pre>)
  }
  return nodes.length > 0 ? <div>{nodes}</div> : null
}

/** Toolview for one excel_* tool: spreadsheet grid, repair diff, or summary. */
export function ExcelToolView(props: ToolCallViewProps): ReactNode {
  const { toolName, block, inputActions } = props
  const [savedNote, setSavedNote] = useState<{ kind: 'success' | 'error'; text: string } | null>(null)
  const [dataVersion, setDataVersion] = useState(0)
  const text = settledText(block)
  /**
   * Parse once per result, not once per render.
   *
   * `JSON.parse` in the render body handed the grid a brand-new `sheets` array
   * on every render of the conversation. The grid's effect keys off that array,
   * so each pass tore the spreadsheet down and rebuilt it; the rebuild forces
   * layout against an emptied card, which resets the surrounding scroll
   * container — the output jumped to the top and then back down, over and over.
   */
  const parsed = useMemo<unknown>(() => {
    if (text === null) return null
    try {
      return JSON.parse(text)
    } catch {
      // Plain-text result (e.g. running fragment) — render as-is.
      return null
    }
  }, [text])
  if (text === null) return null
  const canEdit = TABLE_TOOLS.has(toolName) && typeof inputActions?.setDraft === 'function'
  if (TABLE_TOOLS.has(toolName) && parsed !== null) {
    const value = parsed as TableShape
    if (Array.isArray(value.sheets) && value.sheets.length > 0) {
      const record = parsed as { path?: string; previewPath?: string }
      const filePath = record.path ?? (typeof record.previewPath === 'string' ? record.previewPath.replace(/\.preview\.html$/i, '.xlsx') : undefined)
      const inPlace = Boolean(filePath && commandsRemote)
      const commitEdit = async (cellId: string, newValue: string) => {
        if (!inPlace || !filePath) {
          if (!newValue) return
          const instruction = `请把 ${cellId} 的值改成 ${newValue}，用 excel_operate 的 set 执行，完成后用 excel_read 重新预览并回复“完成”。`
          inputActions.setDraft(instruction)
          inputActions.submit()
          return
        }
        try {
          const outcome = await commandsRemote!.execute(
            String(props.sessionId),
            `/excel-set ${JSON.stringify({ path: filePath, cell: cellId, value: newValue })}`,
          )
          const result = outcome.result
          if (result.kind === 'success') {
            setSavedNote({ kind: 'success', text: `已就地保存：${result.text ?? `${cellId} = ${newValue}`}` })
          } else {
            setSavedNote({ kind: 'error', text: `保存失败：${result.text ?? '未知错误'}` })
          }
        } catch (error) {
          setSavedNote({ kind: 'error', text: `保存失败：${error instanceof Error ? error.message : String(error)}` })
        }
      }
      const rollback = async () => {
        if (!filePath || !commandsRemote) return
        try {
          const outcome = await commandsRemote!.execute(String(props.sessionId), `/excel-undo ${JSON.stringify({ path: filePath })}`)
          const result = outcome.result
          if (result.kind === 'success') {
            setDataVersion((version) => version + 1)
            setSavedNote({ kind: 'success', text: `已回滚：${result.text ?? ''}` })
          } else {
            setSavedNote({ kind: 'error', text: `回滚失败：${result.text ?? '未知错误'}` })
          }
        } catch (error) {
          setSavedNote({ kind: 'error', text: `回滚失败：${error instanceof Error ? error.message : String(error)}` })
        }
      }
      return (
        <div>
          <SpreadsheetView
            key={dataVersion}
            sheets={value.sheets}
            editable={canEdit || inPlace}
            onCellEdited={commitEdit}
          />
          {savedNote !== null && (
            <div style={{ fontSize: 11, color: savedNote.kind === 'success' ? '#15803d' : '#b91c1c', marginTop: 6 }}>
              {savedNote.text}
            </div>
          )}
          {(canEdit || inPlace) && (
            <div style={{ fontSize: 11, color: '#6b7280', marginTop: 6 }}>
              {inPlace
                ? '改完回车 = 就地保存本地文件（自动备份 .bak，可撤销）'
                : '改完回车：变更会通过对话让 agent 执行 set 并重新预览'}
              {inPlace && (
                <button
                  type="button"
                  onClick={() => { void rollback() }}
                  style={{ marginLeft: 10, fontSize: 11, cursor: 'pointer' }}
                >
                  撤销本次修改
                </button>
              )}
            </div>
          )}
        </div>
      )
    }
    if (typeof value.markdown === 'string') {
      return <pre style={{ fontSize: 12, overflow: 'auto', whiteSpace: 'pre-wrap', maxHeight: 720 }}>{value.markdown}</pre>
    }
  }
  if (DIFF_TOOLS.has(toolName) && parsed !== null && typeof parsed === 'object') {
    const diff = DiffFromResult(parsed)
    if (diff !== null) return diff
  }
  if (toolName === 'excel_menu' && parsed !== null && typeof parsed === 'object') {
    const menu = MenuFromResult(parsed)
    if (menu !== null) return menu
  }
  if (toolName === 'excel_insight' && parsed !== null && typeof parsed === 'object') {
    const insight = InsightFromResult(parsed)
    if (insight !== null) return insight
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
      return <pre style={{ fontSize: 12, overflow: 'auto', whiteSpace: 'pre-wrap', maxHeight: 720 }}>{[...lines, ...counts].join('\n')}</pre>
    }
  }
  return <pre style={{ fontSize: 12, overflow: 'auto', whiteSpace: 'pre-wrap', maxHeight: 720 }}>{text}</pre>
}

/** Client plugin body: register the excel_* toolviews into the details column. */
export function apply(ctx: ClientContext): void {
  commandsRemote = (ctx.get('remote.commands') as CommandsRemote | undefined) ?? null
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
