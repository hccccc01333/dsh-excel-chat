import type { ExcelOperation } from './operations.ts'
import type { PlanStep } from './agent.ts'

const CELL_OR_RANGE = /^[A-Za-z]{1,3}\d+$|^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$/

const REQUIRED_STRINGS: Record<string, string[]> = {
  fill: ['source', 'target'],
  fillSeries: ['start', 'target'],
  copyRange: ['source', 'target'],
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
  subtotal: ['sheet', 'range'],
  aggregateReport: ['source'],
  report: ['source'],
  preset: ['source'],
  highlightRows: ['sheet', 'range'],
  fuzzyMatch: ['source', 'target', 'valueColumn', 'outputColumn'],
  findReplace: ['find', 'replace'],
  addSheet: ['name'],
  deleteSheet: ['name'],
  renameSheet: ['oldName', 'newName'],
  duplicateSheet: ['name', 'newName'],
  mailMerge: ['template', 'data'],
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
