import type { LlmText } from './advisor.ts'
import type { AgentPlanContext, AgentPlanner, AgentVerifierContext, PlanStep } from './agent.ts'
import type { ExcelOperation } from './operations.ts'

const OPERATION_CATALOG = [
  'set（写值/公式）', 'fill', 'fillSeries', 'insertRows/deleteRows/insertColumns/deleteColumns',
  'sortRange', 'copyRange', 'style', 'dataValidation', 'conditionalFormatting', 'autoFilter', 'addTable',
  'dedupeRows（去重）', 'fillMissing（补空）', 'removeEmptyRows/removeEmptyColumns', 'trimText',
  'changeCase', 'normalizeText（全角半角）', 'splitColumn（分列）', 'highlightRows（整行高亮）',
  'fuzzyMatch（模糊匹配）', 'subtotal（分类汇总）', 'aggregateReport（动态透视）', 'report（一键报表）',
  'preset（岗位模板）', 'filterToRange', 'findReplace', 'mailMerge', 'addSheet/renameSheet/deleteSheet/duplicateSheet',
  'merge/unmerge', 'freezePanes',
].join('、')

const PARAM_REFERENCE = [
  'set: {"op":"set","cells":{"订单!A1":"值"}}',
  'fillMissing: {"op":"fillMissing","range":"订单!A2:B4","mode":"value|forward|left","value":0}',
  'dedupeRows: {"op":"dedupeRows","sheet":"订单","columns":["A"],"keep":"first|last"}',
  'style: {"op":"style","range":"订单!A1:B1","style":{"bold":true,"fill":"FFFF00","numberFormat":"#,##0.00"}}',
  'sortRange: {"op":"sortRange","range":"订单!A1:B4","keys":[{"column":"B","direction":"asc|desc"}],"headerRows":1}',
  'filterToRange: {"op":"filterToRange","source":"订单!A1:C4","criteria":[{"column":"A","operator":"eq","value":"华东"}],"target":"华东!A1"}',
  'aggregateReport: {"op":"aggregateReport","source":"订单!A1:C4","groupColumn":"A","metrics":[{"column":"C","function":"sum"}],"outputSheet":"汇总"}',
  'report: {"op":"report","source":"订单!A1:B4","groupColumn":"A","metrics":[{"column":"B","function":"sum"}],"outputSheet":"经营报表"}',
  'subtotal: {"op":"subtotal","sheet":"订单","range":"订单!A1:B4","groupColumn":"A","summaryColumns":[{"column":"B","function":"sum"}]}',
  'trimText/changeCase/normalizeText: {"op":"trimText","range":"订单!A2:A4"}',
  'splitColumn: {"op":"splitColumn","sheet":"订单","column":"A","delimiter":"-","startRow":2,"endRow":4}',
  'copyRange: {"op":"copyRange","source":"订单!A2:B3","target":"订单!D2"}',
  'merge: {"op":"merge","range":"订单!A1:B1"}',
  'highlightRows: {"op":"highlightRows","sheet":"订单","range":"订单!A1:B4","criteria":[{"column":"A","operator":"eq","value":"苹果"}]}',
].join('\n')

function stripFence(text: string): string {
  const match = /```(?:json)?\s*([\s\S]*?)```/.exec(text)
  return match ? match[1]!.trim() : text.trim()
}

/**
 * LLM-backed planner/verifier for the goal-driven agent loop. The planner
 * proposes excel_operate steps from the goal + file profile; the verifier
 * judges whether the executed result achieved the goal.
 */
export function createLlmPlanner(llm: LlmText): AgentPlanner {
  return {
    async plan(context: AgentPlanContext): Promise<PlanStep[]> {
      const prompt = [
        '你是 Excel 自动化规划器。你的任务是给下面的目标设计 excel_operate 操作步骤。',
        `用户目标：${context.goal}`,
        `当前文件：${context.path}`,
        `第 ${context.round} 轮。`,
        `工作表：${context.sheetNames.join('、')}`,
        `文件概览：${context.profileSummary}`,
        `公式校验：${context.validationSummary}`,
        ...(context.previousPlan
          ? [
              `上一轮计划：${JSON.stringify(context.previousPlan)}`,
              `上一轮结果：${JSON.stringify(context.previousResult)}`,
              `验证结论：${context.verifierNote ?? ''}`,
              '如果上一轮未达成目标，请根据验证结论修正计划（可追加或修改步骤）。',
            ]
          : []),
        `可用操作：${OPERATION_CATALOG}`,
        `常用操作参数速查：\n${PARAM_REFERENCE}`,
        '返回 ONLY JSON，格式：{"steps":[{"name":"步骤名","operations":[{"op":"操作名",...参数}]}]}。',
        'operations 里的每个对象是 excel_operate 的一个操作，参数按该操作的字段写。',
        '重要：所有 range/source/target/单元格引用必须带工作表前缀（如 "订单!A2:B4"）；需要 sheet 字段的操作必须写 sheet。',
        '必填字段提醒：set 的 cells 是对象；sortRange/filterToRange/aggregateReport/report/preset/subtotal 必须带 keys/metrics/summaryColumns/criteria 数组；fill/fillSeries/copyRange 必须给 source/start/target；renameSheet 用 oldName/newName；addSheet/deleteSheet/duplicateSheet 用 name（duplicateSheet 另给 newName）。',
        '分析类操作：aggregateReport/report/preset 的 metrics 是 [{"column":"C","function":"sum|average|count|counta|max|min"}]，groupColumn 是列字母（如 "A"），outputSheet 给一个表名；subtotal 的 summaryColumns 同理。',
      ].join('\n')
      const text = await llm(prompt)
      const parsed = JSON.parse(stripFence(text)) as { steps?: PlanStep[] }
      if (!Array.isArray(parsed.steps) || parsed.steps.length === 0) {
        throw new Error('planner reply must contain a non-empty steps array')
      }
      const firstSheet = context.sheetNames[0] ?? 'Sheet1'
      return parsed.steps.map((step) => ({
        name: step.name,
        operations: step.operations.map((operation) => normalizeOperation(operation, firstSheet)),
      }))
    },
    async verify(context: AgentVerifierContext): Promise<{ achieved: boolean; reason: string }> {
      const prompt = [
        '你是 Excel 任务验证器。判断用户目标是否已经被执行结果达成。',
        `用户目标：${context.goal}`,
        `执行后的文件概览：${context.profileSummary}`,
        `公式校验：${context.validationSummary}`,
        `本轮计划：${JSON.stringify(context.executedPlan)}`,
        `执行结果摘要：${JSON.stringify(context.executedResult)}`,
        `执行后的单元格快照（前若干非空单元格）：\n${context.cellSnapshot}`,
        '返回 ONLY JSON，格式：{"achieved":true或false,"reason":"简短结论（中文）"}。',
        '必须以单元格快照中的具体证据为准：目标要求的值/公式/样式/结构能在快照中看到才算达成；',
        '公式异常未清零、关键结果缺失、或快照中没有对应证据时一律返回 achieved:false 并说明缺什么。',
        '如果本轮计划执行后文件几乎没变（快照与目标明显不符），必须返回 false。',
        '验证步骤：先把用户目标拆成可检查点（每条要求一条），再逐条对照快照找证据；',
        '每一条都要有明确证据才可返回 achieved:true，任何一条缺证据就返回 false 并指出缺哪条。',
        '例如目标“按区域汇总金额”→ 证据应是区域字段、金额字段、SUMIFS/汇总公式和汇总表都在快照中可见。',
      ].join('\n')
      const text = await llm(prompt)
      const parsed = JSON.parse(stripFence(text)) as { achieved?: unknown; reason?: unknown }
      return { achieved: Boolean(parsed.achieved), reason: String(parsed.reason ?? '') }
    },
  }
}

/** Tolerate planner sloppiness: prefix sheet names onto bare cells/ranges. */
function normalizeOperation(operation: ExcelOperation, firstSheet: string): ExcelOperation {
  const prefix = (value: unknown): unknown => {
    if (typeof value !== 'string' || value.includes('!')) return value
    if (
      /^[A-Za-z]{1,3}\d+$/.test(value) ||
      /^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$/.test(value) ||
      /^[A-Za-z]{1,3}:\d+$/.test(value)
    ) {
      return `${firstSheet}!${value}`
    }
    return value
  }
  const raw = operation as unknown as Record<string, unknown>
  if (operation.op === 'fillMissing') {
    if (raw.mode === undefined) raw.mode = 'value'
    if (raw.value === undefined && raw.fillValue !== undefined) raw.value = raw.fillValue
  }
  if (operation.op === 'renameSheet') {
    if (raw.oldName === undefined && raw.sheet !== undefined) raw.oldName = raw.sheet
    if (raw.newName === undefined && raw.target !== undefined) raw.newName = raw.target
  }
  if (['addSheet', 'deleteSheet', 'hideSheet', 'setTabColor', 'protectSheet', 'unprotectSheet', 'duplicateSheet'].includes(operation.op)) {
    if (raw.name === undefined && raw.sheet !== undefined) raw.name = raw.sheet
    if (operation.op === 'duplicateSheet' && raw.newName === undefined && raw.target !== undefined) raw.newName = raw.target
  }
  if (operation.op === 'filterToRange' && typeof raw.target === 'string' && !raw.target.includes('!') && !/^[A-Za-z]{1,3}\d+$/.test(raw.target)) {
    raw.target = `${raw.target}!A1`
  }
  const out: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(raw)) {
    if (key === 'cells' && value && typeof value === 'object') {
      out[key] = Array.isArray(value)
        ? value.map((id) => prefix(id))
        : Object.fromEntries(
            Object.entries(value as Record<string, unknown>).map(([id, content]) => [prefix(id) as string, String(content)]),
          )
    } else if (key === 'sheet') {
      out[key] = value ?? firstSheet
    } else if (['range', 'source', 'target', 'start'].includes(key)) {
      out[key] = prefix(Array.isArray(value) ? value[0] : value)
    } else {
      out[key] = value
    }
  }
  return out as unknown as ExcelOperation
}
