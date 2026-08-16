import { access, copyFile } from 'node:fs/promises'
import { repairWorkbookFile, type RepairAdvisor } from './repair.ts'
import { writeWorkbookHealthReport } from './health-report.ts'
import type { CellPatch } from './patch.ts'
import type { ValidationResult } from './validator.ts'

export interface AnomalySummary {
  total: number
  byKind: Record<string, number>
  formulaCount: number
  cellCount: number
}

export interface AutofixOutcome {
  repairedPath: string
  repairs: CellPatch[]
  before: AnomalySummary
  after: AnomalySummary
  message: string
  healthScore?: number
  reportSheet?: string
}

/** Collapse a validation result into counts the model can reason about. */
export function summarizeValidation(result: ValidationResult): AnomalySummary {
  const byKind: Record<string, number> = {}
  for (const anomaly of result.anomalies) {
    byKind[anomaly.kind] = (byKind[anomaly.kind] ?? 0) + 1
  }
  return {
    total: result.anomalies.length,
    byKind,
    formulaCount: result.formulaCount,
    cellCount: result.cellCount,
  }
}

/**
 * One-call self-healing loop: validate, apply deterministic repairs (plus an
 * optional LLM advisor), re-validate the repaired copy, and report a compact
 * before/after summary in plain language.
 */
export async function autofixWorkbookFile(
  path: string,
  options: { outPath?: string; advisor?: RepairAdvisor; healthReport?: boolean } = {},
): Promise<AutofixOutcome> {
  const result = await repairWorkbookFile(path, options.advisor, undefined, undefined, options.outPath)
  const includeHealthReport = options.healthReport ?? true
  const before = summarizeValidation(result.before)
  const after = summarizeValidation(result.after)
  const repairedPath = result.repairedPath
  const fixed = result.repairs.length + result.llmRepairs.length
  const lines = [
    `体检：修复前 ${before.total} 个异常，修复后 ${after.total} 个。`,
  ]
  if (fixed > 0) {
    const ids = [...result.repairs, ...result.llmRepairs].map((patch) => patch.id).join(', ')
    lines.push(`已修复 ${fixed} 处：${ids}`)
  } else if (before.total > 0) {
    lines.push(`未自动修复 ${before.total} 处；可提供表格结构后启用 LLM 修复，或人工核对。`)
  } else {
    lines.push('未发现公式异常，无需修复。')
  }
  lines.push(`输出文件：${repairedPath}`)
  const outcome: AutofixOutcome = { repairedPath, repairs: result.repairs, before, after, message: lines.join('\n') }
  if (includeHealthReport) {
    try {
      await access(repairedPath)
    } catch {
      // repairWorkbookFile only writes the copy when repairs exist; keep the
      // "output copy" contract for clean workbooks too.
      await copyFile(path, repairedPath)
    }
    const report = await writeWorkbookHealthReport(repairedPath)
    outcome.healthScore = report.healthScore
    outcome.reportSheet = report.reportSheet
  }
  return outcome
}
