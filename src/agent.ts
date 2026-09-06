import { copyFile, mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import ExcelJS from 'exceljs'
import type { ExcelOperation } from './operations.ts'
import { sanitizeAssertions, sanitizePlan } from './plan-schema.ts'
import { profileWorkbook, type WorkbookProfile } from './profile.ts'
import { buildWorkbookSemanticProfile } from './semantic.ts'
import { runExcelTask, type TaskResult } from './task.ts'
import { verifyWorkbookAssertions, type WorkbookAssertion, type WorkbookVerification } from './verifier.ts'
import { readWorkbookCells, stripPivotTableParts, validateWorkbookFile } from './workbook.ts'

export interface PlanStep {
  name?: string
  operations: ExcelOperation[]
}

/**
 * Verifier 2.0: the planner may accompany its steps with machine-checkable
 * assertions ("expect") that are verified deterministically against the
 * executed workbook, in addition to the LLM verdict.
 */
export interface PlanWithAssertions {
  steps: PlanStep[]
  assertions?: WorkbookAssertion[]
}

export type PlannerPlanOutput = PlanStep[] | PlanWithAssertions

export interface AgentPlanContext {
  goal: string
  path: string
  round: number
  sheetNames: string[]
  profileSummary: string
  semanticSummary?: string
  validationSummary: string
  previousPlan?: PlanStep[]
  previousResult?: TaskResult
  verifierNote?: string
}

export interface AgentVerifierContext extends AgentPlanContext {
  /** Plan that was just executed; its output is at `path`. */
  executedPlan: PlanStep[]
  executedResult: TaskResult
  /** Compact post-execution cell snapshot for evidence-based verification. */
  cellSnapshot: string
}

export interface AgentPlanner {
  plan(context: AgentPlanContext): Promise<PlannerPlanOutput>
  verify(context: AgentVerifierContext): Promise<{ achieved: boolean; reason: string }>
}

export interface AgentRoundResult {
  round: number
  plan: PlanStep[]
  result: TaskResult
  verdict: { achieved: boolean; reason: string }
  deterministicVerification?: WorkbookVerification
  /** Verifier 2.0: deterministic check of the planner's own assertions. */
  planAssertions?: WorkbookVerification
}

export interface AgentTaskResult {
  outputPath: string
  rounds: AgentRoundResult[]
  achieved: boolean
  finalAnomalies: number
}

/**
 * Goal-driven agent loop (Plan -> Act -> Observe -> Verify -> Replan):
 * the planner proposes operation steps for the goal, `runExcelTask` executes
 * them with per-step formula verification and deterministic repair, an LLM
 * verifier checks whether the goal is achieved, and the loop replans up to
 * maxRounds times when it is not.
 */
export async function runAgentTask(
  path: string,
  options: {
    goal: string
    planner: AgentPlanner
    maxRounds?: number
    outPath?: string
    /** Optional hard assertions used by deterministic benchmark/replay callers. */
    deterministicAssertions?: WorkbookAssertion[]
  },
): Promise<AgentTaskResult> {
  const maxRounds = options.maxRounds ?? 2
  if (maxRounds < 1) throw new Error('maxRounds must be at least 1')
  const dir = await mkdtemp(join(tmpdir(), 'vera-agent-'))
  let currentPath = path
  const rounds: AgentRoundResult[] = []
  let achieved = false
  let previousPlan: PlanStep[] | undefined
  let previousResult: TaskResult | undefined
  let verifierNote: string | undefined

  for (let round = 1; round <= maxRounds; round++) {
    const beforeProfile = await profileWorkbook(currentPath)
    const semanticProfile = await buildWorkbookSemanticProfile(currentPath)
    const beforeValidation = await validateWorkbookFile(currentPath)
    const beforeFingerprint = await workbookFingerprint(currentPath)
    const planContext: AgentPlanContext = {
      goal: options.goal,
      path: currentPath,
      round,
      sheetNames: beforeProfile.sheets.map((sheet) => sheet.sheet),
      profileSummary: summarizeProfile(beforeProfile),
      semanticSummary: semanticProfile.summary,
      validationSummary: `${beforeValidation.anomalies.length} 个公式异常`,
      previousPlan,
      previousResult,
      verifierNote,
    }
    let plan: PlanStep[]
    let planAssertions: WorkbookAssertion[] = []
    try {
      const rawPlan = await options.planner.plan(planContext)
      const steps = Array.isArray(rawPlan) ? rawPlan : rawPlan.steps
      if (!steps || steps.length === 0) throw new Error('planner returned an empty plan')
      plan = sanitizePlan(steps, planContext.sheetNames).steps
      if (!Array.isArray(rawPlan) && rawPlan.assertions !== undefined) {
        const sanitized = sanitizeAssertions(rawPlan.assertions, planContext.sheetNames)
        planAssertions = sanitized.assertions
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      if (round < maxRounds) {
        verifierNote = `计划无效：${message}。请修正后重新规划。`
        previousPlan = undefined
        previousResult = undefined
        continue
      }
      throw new Error(`${message}（第 ${round} 轮计划：${summarizePlanOps(plan ?? [])}）`)
    }
    const roundOut = join(dir, `round-${round}.xlsx`)
    let result: TaskResult
    try {
      result = await runExcelTask(currentPath, plan, roundOut)
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      if (round < maxRounds) {
        verifierNote = `执行出错：${message}。请修正计划后重新规划。`
        previousPlan = plan
        previousResult = undefined
        continue
      }
      throw new Error(`${message}（第 ${round} 轮计划：${summarizePlanOps(plan)}）`)
    }
    const afterProfile = await profileWorkbook(result.outputPath)
    const afterValidation = await validateWorkbookFile(result.outputPath)
    const cellSnapshot = await cellSnapshotOf(result.outputPath)
    const changed = (await workbookFingerprint(result.outputPath)) !== beforeFingerprint
    const deterministicVerification = options.deterministicAssertions === undefined
      ? undefined
      : await verifyWorkbookAssertions(result.outputPath, options.deterministicAssertions)
    // Verifier 2.0: deterministic check of the planner's own assertions.
    const assertionCheck = planAssertions.length > 0
      ? await verifyWorkbookAssertions(result.outputPath, planAssertions)
      : undefined
    let verdict = deterministicVerification === undefined
      ? await options.planner.verify({
          ...planContext,
          path: result.outputPath,
          profileSummary: summarizeProfile(afterProfile),
          validationSummary: `${afterValidation.anomalies.length} 个公式异常`,
          executedPlan: plan,
          executedResult: result,
          cellSnapshot,
        })
      : { achieved: deterministicVerification.achieved, reason: deterministicVerification.reason }
    const deterministicNote = `${afterValidation.anomalies.length === 0 ? '公式无异常' : `仍有 ${afterValidation.anomalies.length} 个公式异常`}；文件${changed ? '有' : '没有'}实质变化`
    if (!changed || afterValidation.anomalies.length > 0) {
      verdict = { achieved: false, reason: `${verdict.reason}（确定性校验：${deterministicNote}）` }
    }
    if (assertionCheck && !assertionCheck.achieved) {
      verdict = {
        achieved: false,
        reason: `${verdict.reason}（规划器断言未过 ${assertionCheck.passed}/${assertionCheck.total}：${assertionCheck.failures.slice(0, 3).join('；')}）`,
      }
    }
    rounds.push({ round, plan, result, verdict, deterministicVerification, planAssertions: assertionCheck })
    currentPath = result.outputPath
    previousPlan = plan
    previousResult = result
    verifierNote = verdict.reason
    if (verdict.achieved) {
      achieved = true
      break
    }
  }

  const finalOutput = options.outPath ?? path.replace(/\.xlsx$/i, '.agent.xlsx')
  await copyFile(currentPath, finalOutput)
  const finalAnomalies = (await validateWorkbookFile(finalOutput)).anomalies.length
  return { outputPath: finalOutput, rounds, achieved, finalAnomalies }
}

function summarizeProfile(profile: WorkbookProfile): string {
  return profile.sheets.map((sheet) => {
    const headers = sheet.columns.filter((column) => column.header).map((column) => column.header).slice(0, 8).join(' / ')
    return `${sheet.sheet}：${sheet.dataRows} 行 × ${sheet.columnCount} 列${headers ? `，表头 ${headers}` : ''}`
  }).join('；')
}

function summarizePlanOps(steps: PlanStep[]): string {
  const parts: string[] = []
  for (const step of steps) {
    for (const operation of step.operations) {
      const record = operation as Record<string, unknown>
      const keys = ['range', 'source', 'target', 'start', 'sheet', 'column', 'groupColumn', 'metrics', 'summaryColumns', 'keys', 'criteria', 'filter', 'cells']
      const args = keys
        .filter((key) => record[key] !== undefined)
        .map((key) => `${key}=${JSON.stringify(record[key])}`)
      parts.push(`${operation.op}${args.length > 0 ? `(${args.join(',')})` : ''}`)
    }
  }
  return parts.join(' -> ')
}

/**
 * Compact cell snapshot for LLM verification. Per-sheet round-robin sampling:
 * each sheet (especially summary sheets the plan just created) contributes
 * cells, instead of "first 80 cells of sheet 1" which hides the evidence the
 * verifier needs for analysis tasks.
 */
async function cellSnapshotOf(path: string, limit = 96): Promise<string> {
  const cells = await readWorkbookCells(await readFile(path))
  const bySheet = new Map<string, Array<[string, string]>>()
  for (const [id, content] of Object.entries(cells)) {
    const sheet = id.slice(0, Math.max(0, id.lastIndexOf('!')))
    const list = bySheet.get(sheet) ?? []
    list.push([id, content])
    bySheet.set(sheet, list)
  }
  const groups = [...bySheet.values()]
  const lines: string[] = []
  let index = 0
  // Round-robin across sheets so late-created output sheets are visible.
  while (lines.length < limit && groups.some((group) => index < group.length)) {
    for (const group of groups) {
      const entry = group[index]
      if (!entry) continue
      lines.push(`${entry[0]}=${entry[1].slice(0, 60)}`)
      if (lines.length >= limit) break
    }
    index++
  }
  return lines.join('\n')
}

async function workbookFingerprint(path: string): Promise<string> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(stripPivotTableParts(await readFile(path)) as any)
  const parts: string[] = []
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const fill = cell.fill?.type === 'pattern' ? `|fill=${String((cell.fill.fgColor as { argb?: string } | undefined)?.argb ?? '')}` : ''
        const bold = cell.font?.bold ? '|bold' : ''
        const numFmt = cell.numFmt && cell.numFmt !== 'General' ? `|fmt=${cell.numFmt}` : ''
        const value = cell.formula ? `=${cell.formula}` : cell.value instanceof Date ? cell.value.toISOString() : String(cell.value ?? '')
        parts.push(`${sheet.name}!${cell.address}=${value}${bold}${numFmt}${fill}`)
      })
    })
  })
  return parts.sort().join('|')
}
