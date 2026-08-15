import { copyFile, mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ExcelOperation } from './operations.ts'
import { profileWorkbook, type WorkbookProfile } from './profile.ts'
import { runExcelTask, type TaskResult } from './task.ts'
import { readWorkbookCells, validateWorkbookFile } from './workbook.ts'

export interface PlanStep {
  name?: string
  operations: ExcelOperation[]
}

export interface AgentPlanContext {
  goal: string
  path: string
  round: number
  sheetNames: string[]
  profileSummary: string
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
  plan(context: AgentPlanContext): Promise<PlanStep[]>
  verify(context: AgentVerifierContext): Promise<{ achieved: boolean; reason: string }>
}

export interface AgentRoundResult {
  round: number
  plan: PlanStep[]
  result: TaskResult
  verdict: { achieved: boolean; reason: string }
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
  options: { goal: string; planner: AgentPlanner; maxRounds?: number; outPath?: string },
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
    const beforeValidation = await validateWorkbookFile(currentPath)
    const planContext: AgentPlanContext = {
      goal: options.goal,
      path: currentPath,
      round,
      sheetNames: beforeProfile.sheets.map((sheet) => sheet.sheet),
      profileSummary: summarizeProfile(beforeProfile),
      validationSummary: `${beforeValidation.anomalies.length} 个公式异常`,
      previousPlan,
      previousResult,
      verifierNote,
    }
    const plan = await options.planner.plan(planContext)
    if (!plan || plan.length === 0) throw new Error('planner returned an empty plan')
    const roundOut = join(dir, `round-${round}.xlsx`)
    const result = await runExcelTask(currentPath, plan, roundOut)
    const afterProfile = await profileWorkbook(result.outputPath)
    const afterValidation = await validateWorkbookFile(result.outputPath)
    const cellSnapshot = await cellSnapshotOf(result.outputPath)
    const verdict = await options.planner.verify({
      ...planContext,
      path: result.outputPath,
      profileSummary: summarizeProfile(afterProfile),
      validationSummary: `${afterValidation.anomalies.length} 个公式异常`,
      executedPlan: plan,
      executedResult: result,
      cellSnapshot,
    })
    rounds.push({ round, plan, result, verdict })
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

async function cellSnapshotOf(path: string, limit = 80): Promise<string> {
  const cells = await readWorkbookCells(await readFile(path))
  return Object.entries(cells)
    .slice(0, limit)
    .map(([id, content]) => `${id}=${content.slice(0, 60)}`)
    .join('\n')
}
