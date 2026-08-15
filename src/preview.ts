import { readFile, writeFile } from 'node:fs/promises'
import ExcelJS from 'exceljs'
import { columnToNumber, numberToColumn } from './formula.ts'
import { profileWorkbook } from './profile.ts'
import { stripPivotTableParts } from './workbook.ts'

export interface PreviewOptions {
  sheet?: string
  range?: string
  maxRows?: number
}

export interface PreviewResult {
  markdown: string
  previewPath: string
  summary: string
}

/**
 * Human-facing table preview: render the requested range as a Markdown table
 * (shown inline in the conversation) and write an HTML preview file next to
 * the workbook. Pure ExcelJS + string building, no new dependencies.
 */
export async function buildWorkbookPreview(path: string, options: PreviewOptions = {}): Promise<PreviewResult> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(stripPivotTableParts(await readFile(path)) as any)
  const profile = await profileWorkbook(path, options.sheet)
  const primary = profile.sheets[0]
  const sheet = workbook.worksheets.find((entry) => entry.name.toLowerCase() === (options.sheet ?? primary?.sheet ?? '').toLowerCase())
    ?? workbook.worksheets[0]
  if (!sheet) throw new Error(`no worksheet found in ${path}`)

  const maxRows = options.maxRows ?? 20
  let startCol = 1
  let startRow = 1
  let endCol = sheet.columnCount || 1
  let endRow = sheet.rowCount || 1
  if (options.range) {
    const parsed = parseRange(options.range)
    startCol = parsed.startCol
    startRow = parsed.startRow
    endCol = parsed.endCol
    endRow = parsed.endRow
  }
  endRow = Math.min(endRow, startRow + maxRows - 1)

  const header: string[] = []
  for (let col = startCol; col <= endCol; col++) header.push(numberToColumn(col))
  const rows: string[][] = []
  for (let row = startRow; row <= endRow; row++) {
    const cells: string[] = []
    for (let col = startCol; col <= endCol; col++) {
      cells.push(cellText(sheet.getCell(`${numberToColumn(col)}${row}`)))
    }
    if (cells.some((value) => value !== '')) rows.push(cells)
  }

  const markdown = renderMarkdown(header, rows)
  const previewPath = path.replace(/\.xlsx$/i, '.preview.html')
  await writeFile(previewPath, renderHtml(sheet.name, header, rows), 'utf8')
  const summary = `表 ${sheet.name}：展示 ${rows.length} 行 × ${header.length} 列${options.range ? `（范围 ${options.range}）` : ''}，HTML 预览：${previewPath}`
  return { markdown, previewPath, summary }
}

function parseRange(range: string): { startCol: number; startRow: number; endCol: number; endRow: number } {
  const body = range.includes('!') ? range.slice(range.lastIndexOf('!') + 1) : range
  const match = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(body)
  if (!match) throw new Error(`invalid range: ${range}`)
  return {
    startCol: columnToNumber(match[1]!),
    startRow: Number(match[2]!),
    endCol: match[3] ? columnToNumber(match[3]!) : columnToNumber(match[1]!),
    endRow: match[4] ? Number(match[4]!) : Number(match[2]!),
  }
}

function cellText(cell: ExcelJS.Cell): string {
  if (cell.formula) return `=${cell.formula}`
  const value = cell.value
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return value.toISOString().slice(0, 10)
  if (typeof value === 'object') {
    const text = (value as { text?: unknown }).text
    return typeof text === 'string' ? text : JSON.stringify(value)
  }
  return String(value)
}

function renderMarkdown(header: string[], rows: string[][]): string {
  const lines = [
    `| ${header.join(' | ')} |`,
    `| ${header.map(() => '---').join(' | ')} |`,
    ...rows.map((row) => `| ${row.map(escapeMd).join(' | ')} |`),
  ]
  return lines.join('\n')
}

function escapeMd(value: string): string {
  return value.replace(/\|/g, '\\|').replace(/\n/g, ' ')
}

function renderHtml(sheetName: string, header: string[], rows: string[][]): string {
  const esc = (value: string) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  const head = header.map((value) => `<th>${esc(value)}</th>`).join('')
  const body = rows.map((row) => `<tr>${row.map((value) => `<td>${esc(value)}</td>`).join('')}</tr>`).join('')
  return [
    '<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">',
    `<title>${esc(sheetName)} 预览</title>`,
    '<style>body{font-family:system-ui,sans-serif;margin:24px}table{border-collapse:collapse;width:100%}',
    'th,td{border:1px solid #ddd;padding:6px 10px;text-align:left;font-size:13px}th{background:#f3f4f6;position:sticky;top:0}',
    'tr:nth-child(even){background:#fafafa}</style></head><body>',
    `<h2>${esc(sheetName)} 预览</h2><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`,
    '</body></html>',
  ].join('')
}
