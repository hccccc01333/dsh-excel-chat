/**
 * Failure taxonomy for ExcelBench lite (v0.35): classify each failed LLM
 * task into one explainable category so the benchmark tells us *where* the
 * agent lost points, not just how many tasks failed.
 */

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
    if (/planner|operations 数组|op 字段|sanitize|缺少必填|必填字段|schema|empty plan|必须是对象/i.test(message)) {
      return { category: 'planning', detail: `规划器/计划结构错误：${message.slice(0, 200)}` }
    }
    if (/invalid range|invalid fill|invalid cell|sheet not found|start cell must|必须是数字/i.test(message)) {
      return { category: 'argument', detail: `参数错误：${message.slice(0, 200)}` }
    }
    return { category: 'execution', detail: `执行异常：${message.slice(0, 200)}` }
  }
  if (evidence.verifierFalsePositive) {
    return {
      category: 'verification',
      detail: `验证器判定目标已达成，但断言只过 ${evidence.checksPassed}/${evidence.checksTotal}，完整性异常 ${evidence.integrity}`,
    }
  }
  if (evidence.rounds > 1 && evidence.rounds >= evidence.maxRounds) {
    return {
      category: 'replan',
      detail: `${evidence.rounds} 轮重规划后仍未达成目标，第一轮失败后没有纠正`,
    }
  }
  if (evidence.executedOps.length === 0) {
    return { category: 'planning', detail: '没有执行任何操作' }
  }
  const executed = new Set(evidence.executedOps)
  const missing = evidence.expectedOps.filter((op) => !executed.has(op))
  const unexpected = evidence.executedOps.filter((op) => !evidence.expectedOps.includes(op))
  if (missing.length === 0) {
    if (evidence.argDiffs.length > 0) {
      return { category: 'argument', detail: evidence.argDiffs.slice(0, 3).join('；') }
    }
    return {
      category: 'semantic',
      detail: '期望操作都已执行且参数一致，但断言未过，可能是列/表/指标语义理解偏差',
    }
  }
  if (unexpected.length > 0 && unexpected.every((op) => GENERIC_OPS.has(op))) {
    return {
      category: 'tool-selection',
      detail: `缺少 ${missing.join('/')}，改用通用操作 ${unexpected.join('/')}`,
    }
  }
  if (evidence.expectedOps.some((op) => executed.has(op))) {
    return { category: 'planning', detail: `缺关键步骤：${missing.join('/')}` }
  }
  return {
    category: 'intent',
    detail: `期望操作 ${evidence.expectedOps.join('/')}，实际执行 ${evidence.executedOps.join('/')}`,
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
