const OPERATION_CATALOG = [
    'set（写值/公式）', 'fill', 'fillSeries', 'insertRows/deleteRows/insertColumns/deleteColumns',
    'sortRange', 'copyRange', 'style', 'dataValidation', 'conditionalFormatting', 'autoFilter', 'addTable',
    'dedupeRows（去重）', 'fillMissing（补空）', 'removeEmptyRows/removeEmptyColumns', 'trimText',
    'changeCase', 'normalizeText（全角半角）', 'splitColumn（分列）', 'highlightRows（整行高亮）',
    'fuzzyMatch（模糊匹配）', 'subtotal（分类汇总）', 'aggregateReport（动态透视）', 'report（一键报表）',
    'preset（岗位模板）', 'filterToRange', 'findReplace', 'mailMerge', 'addSheet/renameSheet/deleteSheet/duplicateSheet',
    'merge/unmerge', 'freezePanes',
].join('、');
function stripFence(text) {
    const match = /```(?:json)?\s*([\s\S]*?)```/.exec(text);
    return match ? match[1].trim() : text.trim();
}
/**
 * LLM-backed planner/verifier for the goal-driven agent loop. The planner
 * proposes excel_operate steps from the goal + file profile; the verifier
 * judges whether the executed result achieved the goal.
 */
export function createLlmPlanner(llm) {
    return {
        async plan(context) {
            const prompt = [
                '你是 Excel 自动化规划器。你的任务是给下面的目标设计 excel_operate 操作步骤。',
                `用户目标：${context.goal}`,
                `当前文件：${context.path}`,
                `第 ${context.round} 轮。`,
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
                '返回 ONLY JSON，格式：{"steps":[{"name":"步骤名","operations":[{"op":"操作名",...参数}]}]}。',
                'operations 里的每个对象是 excel_operate 的一个操作，参数按该操作的字段写。',
            ].join('\n');
            const text = await llm(prompt);
            const parsed = JSON.parse(stripFence(text));
            if (!Array.isArray(parsed.steps) || parsed.steps.length === 0) {
                throw new Error('planner reply must contain a non-empty steps array');
            }
            return parsed.steps;
        },
        async verify(context) {
            const prompt = [
                '你是 Excel 任务验证器。判断用户目标是否已经被执行结果达成。',
                `用户目标：${context.goal}`,
                `执行后的文件概览：${context.profileSummary}`,
                `公式校验：${context.validationSummary}`,
                `本轮计划：${JSON.stringify(context.executedPlan)}`,
                `执行结果摘要：${JSON.stringify(context.executedResult)}`,
                '返回 ONLY JSON，格式：{"achieved":true或false,"reason":"简短结论（中文）"}。',
                '只有目标明确达成时才返回 achieved:true；公式异常未清零、关键结果缺失或明显不对时返回 false。',
            ].join('\n');
            const text = await llm(prompt);
            const parsed = JSON.parse(stripFence(text));
            return { achieved: Boolean(parsed.achieved), reason: String(parsed.reason ?? '') };
        },
    };
}
