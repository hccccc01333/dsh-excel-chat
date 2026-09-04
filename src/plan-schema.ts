import type { ExcelOperation } from './operations.ts'
import type { PlanStep } from './agent.ts'

const CELL_OR_RANGE = /^[A-Za-z]{1,3}\d+$|^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$/

const REQUIRED_STRINGS: Record<string, string[]> = {
  fill: ['source', 'target'],
  fillSeries: ['start', 'target'],
  copyRange: ['source', 'target'],
  transpose: ['source', 'target'],
  copyStyle: ['source', 'target'],
  freezeFormulas: ['range'],
  clearRange: ['range'],
  uniqueValues: ['source', 'target'],
  dedupeRows: ['sheet'],
  fillMissing: ['range'],
  removeEmptyRows: ['range'],
  removeEmptyColumns: ['range'],
  trimText: ['range'],
  changeCase: ['range'],
  normalizeText: ['range'],
  splitColumn: ['sheet', 'column'],
  sortRange: ['range'],
  filterToRange: ['source', 'target'],
  style: ['range'],
  merge: ['range'],
  unmerge: ['range'],
  subtotal: ['sheet', 'range', 'groupColumn'],
  aggregateReport: ['source', 'groupColumn'],
  report: ['source', 'groupColumn'],
  preset: ['source', 'groupColumn'],
  highlightRows: ['sheet', 'range'],
  fuzzyMatch: ['source', 'target', 'valueColumn', 'outputColumn'],
  findReplace: ['find', 'replace'],
  addSheet: ['name'],
  deleteSheet: ['name'],
  renameSheet: ['oldName', 'newName'],
  duplicateSheet: ['name', 'newName'],
  mailMerge: ['template', 'data'],
  hideRows: ['sheet'],
  hideColumns: ['sheet'],
  groupRows: ['sheet'],
  groupColumns: ['sheet', 'from', 'to'],
  autoFitColumnWidths: ['sheet'],
  unfreezePanes: ['sheet'],
  unmergeAll: ['sheet'],
  setZoom: ['sheet'],
  showGridLines: ['sheet'],
  headerFooter: ['sheet'],
  printTitles: ['sheet'],
  moveSheet: ['name'],
  setHyperlink: ['cell'],
  joinSheets: ['source', 'sourceKey', 'lookup', 'lookupKey'],
  rankColumn: ['range', 'metricColumn', 'outputColumn'],
  rowPageBreaks: ['sheet'],
  clearPageBreaks: ['sheet'],
  addComment: ['cell', 'text'],
  addSparklines: ['dataRange', 'locationRange'],
}

const REQUIRED_ARRAYS: Record<string, string[]> = {
  sortRange: ['keys'],
  filterToRange: ['criteria'],
  aggregateReport: ['metrics'],
  report: ['metrics'],
  preset: ['metrics'],
  subtotal: ['summaryColumns'],
  highlightRows: ['criteria'],
  conditionalFormatting: ['rules'],
  joinSheets: ['valueColumns', 'outputColumns'],
  hideColumns: ['columns'],
  rowPageBreaks: ['rows'],
}

const REQUIRED_NUMBERS: Record<string, string[]> = {
  freezePanes: ['row'],
  splitColumn: ['startRow'],
  setColumnWidth: ['width'],
  setRowHeight: ['row', 'height'],
  hideRows: ['from', 'to'],
  groupRows: ['start', 'end'],
  setZoom: ['zoom'],
  moveSheet: ['position'],
}

const REQUIRED_STRING_EXTRA: Record<string, string[]> = {
  freezePanes: ['sheet', 'column'],
}

export interface SanitizedPlan {
  steps: PlanStep[]
  notes: string[]
}

/**
 * Validate and repair a planner-produced plan before execution. Salvageable
 * issues are fixed in place (sheet prefix, missing sheet, array wrapping,
 * alias fields, cell values); unsalvageable issues throw so the agent loop
 * can feed the exact message back to the planner for a corrected plan.
 */
export function sanitizePlan(steps: PlanStep[], sheetNames: string[]): SanitizedPlan {
  const firstSheet = sheetNames[0] ?? 'Sheet1'
  const notes: string[] = []
  const prefix = (value: unknown): unknown => {
    if (typeof value !== 'string' || value.includes('!')) return value
    return CELL_OR_RANGE.test(value) ? `${firstSheet}!${value}` : value
  }
  const out: PlanStep[] = []
  steps.forEach((step, stepIndex) => {
    if (!Array.isArray(step.operations)) {
      throw new Error(`第 ${stepIndex + 1} 步没有 operations 数组`)
    }
    const operations = step.operations.map((operation, opIndex) => {
      if (!operation || typeof operation.op !== 'string') {
        throw new Error(`第 ${stepIndex + 1} 步第 ${opIndex + 1} 个操作缺少 op 字段`)
      }
      const raw = { ...operation } as Record<string, unknown>
      for (const key of ['range', 'source', 'target', 'start'] as const) {
        const value = Array.isArray(raw[key]) ? raw[key]![0] : raw[key]
        if (value !== undefined && typeof value === 'string' && !value.includes('!') && CELL_OR_RANGE.test(value)) {
          raw[key] = `${firstSheet}!${value}`
          notes.push(`${operation.op} 的 ${key} 已补工作表前缀`)
        }
      }
      if (raw.sheet === undefined && (REQUIRED_STRINGS[operation.op] ?? []).includes('sheet')) {
        raw.sheet = firstSheet
        notes.push(`${operation.op} 已补默认工作表`)
      }
      for (const key of ['sheet', 'column', 'groupColumn', 'valueColumn', 'outputColumn', 'sourceKey', 'targetKey', 'name', 'oldName', 'newName', 'template', 'data', 'find', 'replace', 'delimiter'] as const) {
        const value = raw[key]
        if (value === undefined) continue
        const scalar = Array.isArray(value) ? value[0] : value
        if (scalar !== undefined && typeof scalar !== 'string') {
          raw[key] = String(scalar)
          notes.push(`${operation.op} 的 ${key} 已转为字符串`)
        }
      }
      for (const field of ['metrics', 'summaryColumns', 'keys', 'criteria'] as const) {
        const list = raw[field]
        if (!Array.isArray(list)) continue
        raw[field] = list.map((entry, index) => {
          if (entry === null || typeof entry !== 'object') {
            throw new Error(`${operation.op} 的 ${field}[${index}] 必须是对象`)
          }
          const record = { ...(entry as Record<string, unknown>) }
          if (record.column !== undefined && typeof record.column !== 'string') {
            record.column = String(record.column)
            notes.push(`${operation.op} 的 ${field}[${index}].column 已转为字符串`)
          }
          if (record.column === undefined || record.column === '') {
            throw new Error(`${operation.op} 的 ${field}[${index}].column 缺失`)
          }
          return record
        })
      }
      if (operation.op === 'freezePanes' && raw.row === undefined && typeof raw.range === 'string') {
        const body = raw.range.includes('!') ? raw.range.slice(raw.range.lastIndexOf('!') + 1) : raw.range
        const match = /^([A-Za-z]{1,3})(\d+)$/.exec(body)
        if (match) {
          raw.column = match[1]!
          raw.row = Number(match[2]!)
          notes.push(`freezePanes 已从 range ${raw.range} 推导 row/column`)
        }
      }
      if (operation.op === 'crosstab') {
        // Planners sometimes emit flat metric fields instead of the object.
        if (raw.metric === undefined && (raw.metricColumn !== undefined || raw.metricFunction !== undefined)) {
          raw.metric = { column: raw.metricColumn, function: raw.metricFunction ?? 'sum' }
          notes.push('crosstab 的 metricColumn/metricFunction 已合并为 metric 对象')
        }
        if (raw.metric !== undefined && raw.metric !== null && typeof raw.metric === 'object' && !Array.isArray(raw.metric)) {
          const metric = { ...(raw.metric as Record<string, unknown>) }
          if (metric.function === undefined) {
            metric.function = 'sum'
            notes.push('crosstab 的 metric.function 已补默认 sum')
          }
          if (typeof metric.function !== 'string' || !['sum', 'average', 'count', 'counta', 'max', 'min'].includes(metric.function)) {
            throw new Error(`crosstab 的 metric.function 不支持：${String(metric.function)}`)
          }
          if (metric.column !== undefined && typeof metric.column !== 'string') {
            metric.column = String(metric.column)
            notes.push('crosstab 的 metric.column 已转为字符串')
          }
          raw.metric = metric
        }
      }
      for (const key of REQUIRED_NUMBERS[operation.op] ?? []) {
        const value = raw[key]
        if (value === undefined || value === null || value === '') {
          throw new Error(`${operation.op} 缺少必填数字 ${key}`)
        }
        const number = typeof value === 'string' ? Number(value) : value
        if (typeof number !== 'number' || !Number.isFinite(number)) {
          throw new Error(`${operation.op} 的 ${key} 必须是数字`)
        }
        raw[key] = number
      }
      if (operation.op === 'fill' || operation.op === 'fillSeries') {
        const start = raw.start ?? raw.source
        const target = raw.target
        const targetBody = typeof target === 'string'
          ? (target.includes('!') ? target.slice(target.lastIndexOf('!') + 1) : target)
          : ''
        if (typeof start === 'string' && typeof target === 'string' && !targetBody.includes(':') && CELL_OR_RANGE.test(targetBody)) {
          raw.target = `${start}:${targetBody}`
          notes.push(`${operation.op} 的 target 已扩展为 ${start}:${targetBody}`)
        }
      }
      for (const field of REQUIRED_ARRAYS[operation.op] ?? []) {
        if (raw[field] === undefined) {
          throw new Error(`${operation.op} 缺少必填数组 ${field}`)
        }
        if (!Array.isArray(raw[field])) {
          raw[field] = [raw[field]]
          notes.push(`${operation.op} 的 ${field} 已包装为数组`)
        }
      }
      for (const field of REQUIRED_STRINGS[operation.op] ?? []) {
        if (raw[field] === undefined || raw[field] === null || raw[field] === '') {
          throw new Error(`${operation.op} 缺少必填字段 ${field}`)
        }
      }
      for (const field of REQUIRED_STRING_EXTRA[operation.op] ?? []) {
        if (raw[field] === undefined || raw[field] === null || raw[field] === '') {
          throw new Error(`${operation.op} 缺少必填字段 ${field}`)
        }
      }
      if (raw.cells && typeof raw.cells === 'object') {
        if (Array.isArray(raw.cells)) {
          raw.cells = raw.cells.map((id) => prefix(id))
        } else if (operation.op === 'set') {
          raw.cells = Object.fromEntries(
            Object.entries(raw.cells as Record<string, unknown>).map(([id, content]) => [prefix(id) as string, String(content)]),
          )
        }
      }
      if (operation.op === 'fillMissing') {
        if (raw.mode === undefined) {
          raw.mode = 'value'
          notes.push('fillMissing 已补 mode=value')
        }
        if (raw.value === undefined && raw.fillValue !== undefined) {
          raw.value = raw.fillValue
          notes.push('fillMissing 的 fillValue 已改为 value')
        }
      }
      if (operation.op === 'renameSheet') {
        if (raw.oldName === undefined && raw.sheet !== undefined) raw.oldName = raw.sheet
        if (raw.newName === undefined && raw.target !== undefined) raw.newName = raw.target
      }
      if (['addSheet', 'deleteSheet', 'hideSheet', 'setTabColor', 'protectSheet', 'unprotectSheet', 'duplicateSheet'].includes(operation.op)) {
        if (raw.name === undefined && raw.sheet !== undefined) raw.name = raw.sheet
        if (operation.op === 'duplicateSheet' && raw.newName === undefined && raw.target !== undefined) raw.newName = raw.target
      }
      if (operation.op === 'filterToRange' && typeof raw.target === 'string' && !raw.target.includes('!') && !CELL_OR_RANGE.test(raw.target)) {
        raw.target = `${raw.target}!A1`
        notes.push('filterToRange 的 target 已补 !A1')
      }
      return raw as unknown as ExcelOperation
    })
    out.push({ name: step.name, operations })
  })
  return { steps: out, notes }
}
