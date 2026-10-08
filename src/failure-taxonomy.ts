/**
 * Failure taxonomy for ExcelBench lite (v0.35): classify each failed LLM
 * task into one explainable category so the benchmark tells us *where* the
 * agent lost points, not just how many tasks failed.
 */

import { t, listJoin, listSeparator } from './i18n.ts'
export type FailureCategory =
  | 'intent'
  | 'semantic'
  | 'planning'
  | 'tool-selection'
  | 'argument'
  | 'execution'
  | 'verification'
  | 'replan'
  | 'other'

export const FAILURE_CATEGORY_LABELS: Record<FailureCategory, string> = {
  intent: 'Intent Error',
  semantic: 'Semantic Error',
  planning: 'Planning Error',
  'tool-selection': 'Tool Selection Error',
  argument: 'Argument Error',
  execution: 'Execution Error',
  verification: 'Verification Error',
  replan: 'Replan Error',
  other: 'Other',
}

export interface FailureClassification {
  category: FailureCategory
  detail: string
}

export interface FailureEvidence {
  /** Agent loop crashed before completing a round (thrown error). */
  crashed?: boolean
  error?: string | null
  /**
   * Kind carried by a `PlanSchemaError`. Classifying by this instead of by the
   * message text is what lets the plan-schema messages be translated at all:
   * matching Chinese prose meant an English message silently became an
   * "execution" error.
   */
  errorKind?: 'planning' | 'argument'
  /** Verifier claimed the goal was achieved while final checks failed. */
  verifierFalsePositive?: boolean
  /** Rounds actually executed. */
  rounds: number
  /** Maximum rounds allowed by the loop. */
  maxRounds: number
  /** Operation names the agent actually executed (all rounds). */
  executedOps: string[]
  /** Operation names the canonical plan requires. */
  expectedOps: string[]
  /** Human-readable diffs between executed and expected arguments. */
  argDiffs: string[]
  checksPassed: number
  checksTotal: number
  integrity: number
}

/** Generic fallback operations the model often picks instead of a specific tool. */
const GENERIC_OPS = new Set(['set', 'fill', 'style', 'sortRange', 'clear'])

/** Classify one failed task with deterministic, explainable heuristics. */
export function classifyFailure(evidence: FailureEvidence): FailureClassification {
  if (evidence.crashed) {
    const message = evidence.error ?? ''
    if (evidence.errorKind === 'planning' || /planner|sanitize|schema|empty plan/i.test(message)) {
      return { category: 'planning', detail: t('规划器/计划结构错误：{message}', { message: message.slice(0, 200) }) }
    }
    if (evidence.errorKind === 'argument' || /invalid range|invalid fill|invalid cell|sheet not found|start cell must/i.test(message)) {
      return { category: 'argument', detail: t('参数错误：{message}', { message: message.slice(0, 200) }) }
    }
    return { category: 'execution', detail: t('执行异常：{message}', { message: message.slice(0, 200) }) }
  }
  if (evidence.verifierFalsePositive) {
    return {
      category: 'verification',
      detail: t('验证器判定目标已达成，但断言只过 {passed}/{total}，完整性异常 {integrity}', { passed: evidence.checksPassed, total: evidence.checksTotal, integrity: evidence.integrity }),
    }
  }
  if (evidence.rounds > 1 && evidence.rounds >= evidence.maxRounds) {
    return {
      category: 'replan',
      detail: t('{rounds} 轮重规划后仍未达成目标，第一轮失败后没有纠正', { rounds: evidence.rounds }),
    }
  }
  if (evidence.executedOps.length === 0) {
    return { category: 'planning', detail: t('没有执行任何操作') }
  }
  const executed = new Set(evidence.executedOps)
  const missing = evidence.expectedOps.filter((op) => !executed.has(op))
  const unexpected = evidence.executedOps.filter((op) => !evidence.expectedOps.includes(op))
  if (missing.length === 0) {
    if (evidence.argDiffs.length > 0) {
      return { category: 'argument', detail: listJoin(evidence.argDiffs.slice(0, 3), 'semicolon') }
    }
    return {
      category: 'semantic',
      detail: t('期望操作都已执行且参数一致，但断言未过，可能是列/表/指标语义理解偏差'),
    }
  }
  if (unexpected.length > 0 && unexpected.every((op) => GENERIC_OPS.has(op))) {
    return {
      category: 'tool-selection',
      detail: t('缺少 {missing}，改用通用操作 {unexpected}', { missing: missing.join('/'), unexpected: unexpected.join('/') }),
    }
  }
  if (evidence.expectedOps.some((op) => executed.has(op))) {
    return { category: 'planning', detail: t('缺关键步骤：{missing}', { missing: missing.join('/') }) }
  }
  return {
    category: 'intent',
    detail: t('期望操作 {expected}，实际执行 {actual}', { expected: evidence.expectedOps.join('/'), actual: evidence.executedOps.join('/') }),
  }
}

/** Aggregate a report's failed tasks into category counts, largest first. */
export function summarizeFailureBreakdown(
  failures: Array<{ category: FailureCategory }>,
): Record<FailureCategory, number> {
  const counts = Object.fromEntries(
    (Object.keys(FAILURE_CATEGORY_LABELS) as FailureCategory[]).map((category) => [category, 0]),
  ) as Record<FailureCategory, number>
  for (const failure of failures) counts[failure.category] += 1
  return counts
}
