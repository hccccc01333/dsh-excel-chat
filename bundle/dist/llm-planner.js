const OPERATION_CATALOG = [
    '汇总/透视（活公式，优先用）',
    '  aggregateReport（动态透视汇总表：分组+SUMIFS 活公式+总计）',
    '  crosstab（二维交叉表：行维度×列维度，单元格是 SUMIFS/COUNTIFS 活公式）',
    '  subtotal（在原表插入分类小计/总计行，需先按分组列排序）',
    '  report（一键报表：排序+小计+汇总表+筛选+表头样式+冻结）',
    '  preset（岗位模板 ops/product/data）',
    '  rankColumn（RANK 排名列）、uniqueValues（提取去重值列表）',
    '两表关联',
    '  joinSheets（精确匹配回填多列，无公式 VLOOKUP）',
    '  fuzzyMatch（名称模糊匹配回填）',
    '  filterToRange（高级筛选到目标区域，matchAll=true 为 AND）',
    '数据编辑',
    '  set（写值/公式）、fill（公式填充）、fillSeries（序列）',
    '  copyRange（move:true 移动；valuesOnly:true 只粘贴值）',
    '  transpose（转置粘贴）、copyStyle（格式刷）、freezeFormulas（公式转值）',
    '  insertRows/deleteRows/insertColumns/deleteColumns（引用联动）',
    '  sortRange（排序）、findReplace、merge/unmerge/unmergeAll、clearRange',
    '数据清洗',
    '  dedupeRows（去重）、fillMissing（补空）、removeEmptyRows/removeEmptyColumns',
    '  trimText、changeCase（大小写）、normalizeText（全角半角）、splitColumn（分列）',
    '  highlightRows（整行条件高亮）',
    '样式与版式',
    '  style（字体/填充/边框/对齐/数字格式/删除线/旋转）',
    '  dataValidation、conditionalFormatting（数据条/色阶/图标集/重复值）',
    '  autoFilter、addTable、setColumnWidth、setRowHeight、autoFitColumnWidths（自适应列宽）',
    '  freezePanes/unfreezePanes、hideRows/hideColumns、groupRows/groupColumns（分组折叠）',
    '打印与页面',
    '  pageSetup（打印区域/方向/缩放/页边距/水平垂直居中）',
    '  printTitles（每页重复标题行/列）、headerFooter（页眉页脚 &-code）',
    '  rowPageBreaks（手动分页符）/clearPageBreaks',
    '批注与迷你图',
    '  addComment（单元格批注）、addSparklines（每行趋势迷你图）',
    '超链接与工作簿',
    '  setHyperlink（外部 URL/站内跳转）、addSheet/renameSheet/deleteSheet/duplicateSheet/hideSheet/setTabColor/moveSheet',
    '  setWorkbookProperties（作者/标题/关键词/打开时重算）',
].join('\n');
const PARAM_REFERENCE = [
    'set: {"op":"set","cells":{"订单!A1":"值"}}',
    'aggregateReport: {"op":"aggregateReport","source":"订单!A1:C4","groupColumn":"A","metrics":[{"column":"C","function":"sum"}],"outputSheet":"汇总"}',
    'crosstab: {"op":"crosstab","source":"订单!A1:C9","rowColumn":"A","columnColumn":"B","metric":{"column":"C","function":"sum"},"outputSheet":"交叉表","totals":true}',
    'joinSheets: {"op":"joinSheets","source":"订单!A1:B4","sourceKey":"A","lookup":"客户!A1:C3","lookupKey":"A","valueColumns":["B","C"],"outputColumns":["C","D"],"missValue":"未匹配"}',
    'report: {"op":"report","source":"订单!A1:B4","groupColumn":"A","metrics":[{"column":"B","function":"sum"}],"outputSheet":"经营报表"}',
    'subtotal: {"op":"subtotal","sheet":"订单","range":"订单!A1:B4","groupColumn":"A","summaryColumns":[{"column":"B","function":"sum"}]}',
    'filterToRange: {"op":"filterToRange","source":"订单!A1:C4","criteria":[{"column":"A","operator":"eq","value":"华东"}],"target":"华东!A1"}',
    'preset: {"op":"preset","role":"ops|product|data","source":"订单!A1:C4","groupColumn":"A","metrics":[{"column":"C","function":"sum"}],"filter":{"column":"A","operator":"eq","value":"华东"}}',
    'fuzzyMatch: {"op":"fuzzyMatch","source":"订单!A1:B4","sourceKey":"A","target":"价目表!A1:B4","targetKey":"A","valueColumn":"B","outputColumn":"C","threshold":0.6}',
    'rankColumn: {"op":"rankColumn","range":"订单!A1:B4","metricColumn":"B","outputColumn":"C","descending":true,"skipHeader":true}',
    'uniqueValues: {"op":"uniqueValues","source":"订单!A1:A7","target":"订单!E1","includeHeader":true}',
    'fillMissing: {"op":"fillMissing","range":"订单!A2:B4","mode":"value|forward|left","value":0}',
    'dedupeRows: {"op":"dedupeRows","sheet":"订单","columns":["A"],"keep":"first|last"}',
    'style: {"op":"style","range":"订单!A1:B1","style":{"bold":true,"fill":"FFFF00","numberFormat":"#,##0.00"}}',
    'sortRange: {"op":"sortRange","range":"订单!A1:B4","keys":[{"column":"B","direction":"asc|desc"}],"headerRows":1}',
    'copyRange: {"op":"copyRange","source":"订单!A2:B3","target":"订单!D2","valuesOnly":false}',
    'fill: {"op":"fill","source":"订单!D2","target":"订单!D2:D11"}',
    'fillSeries: {"op":"fillSeries","start":"订单!A2","target":"订单!A2:A11","step":1}',
    'renameSheet: {"op":"renameSheet","oldName":"Sheet1","newName":"订单"}',
    'addSheet/deleteSheet/hideSheet: {"op":"addSheet","name":"汇总"}；duplicateSheet 用 name+newName',
    'clearRange: {"op":"clearRange","range":"订单!D2:E9","mode":"contents|formats|all"}',
    'transpose: {"op":"transpose","source":"订单!A1:C4","target":"转置!A1"}',
    'copyStyle: {"op":"copyStyle","source":"订单!A1","target":"订单!B1:B9"}',
    'freezeFormulas: {"op":"freezeFormulas","range":"订单!D2:D9"}',
    'hideRows: {"op":"hideRows","sheet":"订单","from":2,"to":9,"hidden":true}',
    'hideColumns: {"op":"hideColumns","sheet":"订单","columns":["C"],"hidden":true}',
    'groupRows: {"op":"groupRows","sheet":"订单","start":3,"end":6,"collapse":true}',
    'autoFitColumnWidths: {"op":"autoFitColumnWidths","sheet":"订单"}',
    'freezePanes: {"op":"freezePanes","sheet":"订单","row":2,"column":"A"}',
    'pageSetup: {"op":"pageSetup","sheet":"订单","printArea":"A1:F40","orientation":"landscape","fitToWidth":1}',
    'printTitles: {"op":"printTitles","sheet":"订单","rows":"1:1"}',
    'headerFooter: {"op":"headerFooter","sheet":"订单","oddFooter":"第 &P 页 / 共 &N 页"}',
    'rowPageBreaks: {"op":"rowPageBreaks","sheet":"订单","rows":[11,21]}',
    'addComment: {"op":"addComment","cell":"订单!B2","text":"该行金额待复核","author":"张三"}',
    'addSparklines: {"op":"addSparklines","dataRange":"订单!B2:F31","locationRange":"订单!G2:G31","type":"line"}',
    'setHyperlink: {"op":"setHyperlink","cell":"订单!A1","url":"https://example.com","text":"官网"}',
    'merge: {"op":"merge","range":"订单!A1:B1"}',
    'highlightRows: {"op":"highlightRows","sheet":"订单","range":"订单!A1:B4","criteria":[{"column":"A","operator":"eq","value":"苹果"}]}',
    'splitColumn: {"op":"splitColumn","sheet":"订单","column":"A","delimiter":"-","startRow":2,"endRow":4}',
    'trimText/changeCase/normalizeText: {"op":"trimText","range":"订单!A2:A4"}',
    'moveSheet: {"op":"moveSheet","name":"汇总","position":1}',
].join('\n');
const FEW_SHOT = [
    '示例 1（汇总）——目标“按区域汇总订单金额，输出到汇总表”：',
    '{"steps":[{"name":"动态透视汇总","operations":[{"op":"aggregateReport","source":"订单!A1:C40","groupColumn":"A","metrics":[{"column":"C","function":"sum"}],"outputSheet":"汇总"}]}]}',
    '示例 2（两表关联）——目标“从客户表按客户名把负责人和电话补到订单表”：',
    '{"steps":[{"name":"关联客户表","operations":[{"op":"joinSheets","source":"订单!A1:B40","sourceKey":"A","lookup":"客户!A1:C30","lookupKey":"A","valueColumns":["B","C"],"outputColumns":["C","D"],"missValue":"未匹配"}]}]}',
    '示例 3（交叉表）——目标“做地区×季度的销售额交叉表”：',
    '{"steps":[{"name":"二维交叉表","operations":[{"op":"crosstab","source":"数据!A1:C50","rowColumn":"A","columnColumn":"B","metric":{"column":"C","function":"sum"},"outputSheet":"交叉表","totals":true}]}]}',
    '示例 4（排名+版式）——目标“给分数加名次列，表头加粗冻结并自适应列宽”：',
    '{"steps":[{"name":"排名列","operations":[{"op":"rankColumn","range":"成绩!A1:B31","metricColumn":"B","outputColumn":"C","descending":true}]},{"name":"版式","operations":[{"op":"style","range":"成绩!A1:C1","style":{"bold":true}},{"op":"freezePanes","sheet":"成绩","row":2,"column":"A"},{"op":"autoFitColumnWidths","sheet":"成绩"}]}]}',
].join('\n');
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
                `工作表：${context.sheetNames.join('、')}`,
                `文件概览：${context.profileSummary}`,
                `语义画像：${context.semanticSummary ?? '（无）'}`,
                `公式校验：${context.validationSummary}`,
                ...(context.previousPlan
                    ? [
                        `上一轮计划：${JSON.stringify(context.previousPlan)}`,
                        `上一轮结果：${JSON.stringify(context.previousResult)}`,
                        `验证结论：${context.verifierNote ?? ''}`,
                        '如果上一轮未达成目标，请根据验证结论修正计划（可追加或修改步骤）。',
                    ]
                    : []),
                `可用操作（按用途分组）：\n${OPERATION_CATALOG}`,
                `常用操作参数速查：\n${PARAM_REFERENCE}`,
                `完整示例（照此 JSON 结构输出）：\n${FEW_SHOT}`,
                '返回 ONLY JSON，格式：{"steps":[{"name":"步骤名","operations":[{"op":"操作名",...参数}]}]}。',
                'operations 里的每个对象是 excel_operate 的一个操作，参数按该操作的字段写。',
                '重要：所有 range/source/target/单元格引用必须带工作表前缀（如 "订单!A2:B4"）；需要 sheet 字段的操作必须写 sheet。',
                '- 工作表名必须用「工作表」清单里的精确名称（如清单是"订单"就不要写成"订单表"）；renameSheet 的 oldName 是现有表、newName 才是新名字。',
                '分析类任务规则：',
                '- 汇总/透视/交叉表/排名一律用 aggregateReport/crosstab/subtotal/rankColumn 等活公式操作，禁止用 set 写死汇总数字（验证器会因快照中无公式而判未达成）。',
                '- source/数据区域必须包含表头行且覆盖全部数据行（概览里有行数，如 39 行数据则写到第 40 行）。',
                '- groupColumn/rowColumn/columnColumn 必须是语义画像里的维度列；metrics 的 column 必须是指标列；画像没提到的列不要猜。',
                '- aggregateReport/crosstab 会自动建输出表，不要先 addSheet；outputSheet 起一个新名字避免覆盖。',
                '- crosstab 的 metric 是对象 {"column":"C","function":"sum"}，不是字符串。',
                '- preset 必须带 role（ops 运营 / product 产品 / data 数分，按目标判断）。',
                '- aggregateReport 的 outputSheet 默认命名「汇总」、crosstab 默认「交叉表」，除非用户指定了名字。',
                '- joinSheets 的 valueColumns/outputColumns 是等长数组，一一对应。',
                '- fill 的 source 是单元格、target 是区域；不要用 fill 顶替 set。',
                '- 上一轮计划执行了但验证不通过时，优先补缺失的步骤而不是整体重来。',
            ].join('\n');
            const text = await llm(prompt);
            const parsed = JSON.parse(stripFence(text));
            if (!Array.isArray(parsed.steps) || parsed.steps.length === 0) {
                throw new Error('planner reply must contain a non-empty steps array');
            }
            const firstSheet = context.sheetNames[0] ?? 'Sheet1';
            return parsed.steps.map((step) => ({
                name: step.name,
                operations: step.operations.map((operation) => normalizeOperation(operation, firstSheet)),
            }));
        },
        async verify(context) {
            const prompt = [
                '你是 Excel 任务验证器。判断用户目标是否已经被执行结果达成。',
                `用户目标：${context.goal}`,
                `执行后的文件概览：${context.profileSummary}`,
                `语义画像：${context.semanticSummary ?? '（无）'}`,
                `公式校验：${context.validationSummary}`,
                `本轮计划：${JSON.stringify(context.executedPlan)}`,
                `执行结果摘要：${JSON.stringify(context.executedResult)}`,
                `执行后的单元格快照（按工作表轮询采样，覆盖每个表包括新建的输出表）：\n${context.cellSnapshot}`,
                '返回 ONLY JSON，格式：{"achieved":true或false,"reason":"简短结论（中文）"}。',
                '必须以单元格快照中的具体证据为准：目标要求的值/公式/样式/结构能在快照中看到才算达成；',
                '公式异常未清零、关键结果缺失、或快照中没有对应证据时一律返回 achieved:false 并说明缺什么。',
                '如果本轮计划执行后文件几乎没变（快照与目标明显不符），必须返回 false。',
                '验证步骤：先把用户目标拆成可检查点（每条要求一条），再逐条对照快照找证据；',
                '每一条都要有明确证据才可返回 achieved:true，任何一条缺证据就返回 false 并指出缺哪条。',
                '按目标类型的证据标准：',
                '- 汇总/透视/交叉表：快照必须能看到输出工作表的表头行、至少一个分组行、以及以 =SUMIFS(/=COUNTIFS( 等开头的活公式或聚合数字；只有输入原表没变、没有输出表 → false。',
                '- 新建工作表/改名/排名列/公式：快照中必须有对应工作表名或 = 开头的公式单元格。',
                '- 样式/隐藏/冻结/分组/打印/工作簿属性：快照看不到，必须核对本轮计划 JSON 里确有对应操作（style/hideRows/freezePanes/groupRows/printTitles/headerFooter/setWorkbookProperties…），计划里没有就是没做 → false。',
                '- 数据清洗：快照中重复值/空值/空格应已消失；原样还在 → false。',
                '反例：目标“按区域汇总金额”，本轮计划只用 set 写了几个静态数字、快照里没有任何 =SUMIFS( 公式或输出汇总表 → 必须返回 achieved:false，理由写明“缺少活公式汇总”。',
                '用语义画像核对：目标里的指标/维度必须在画像的指标/维度列中有对应，找不到对应说明任务没做对，返回 false。',
            ].join('\n');
            const text = await llm(prompt);
            const parsed = JSON.parse(stripFence(text));
            return { achieved: Boolean(parsed.achieved), reason: String(parsed.reason ?? '') };
        },
    };
}
/** Tolerate planner sloppiness: prefix sheet names onto bare cells/ranges. */
function normalizeOperation(operation, firstSheet) {
    const prefix = (value) => {
        if (typeof value !== 'string' || value.includes('!'))
            return value;
        if (/^[A-Za-z]{1,3}\d+$/.test(value) ||
            /^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$/.test(value) ||
            /^[A-Za-z]{1,3}:\d+$/.test(value)) {
            return `${firstSheet}!${value}`;
        }
        return value;
    };
    const raw = operation;
    if (operation.op === 'fillMissing') {
        if (raw.mode === undefined)
            raw.mode = 'value';
        if (raw.value === undefined && raw.fillValue !== undefined)
            raw.value = raw.fillValue;
    }
    if (operation.op === 'fillSeries' && raw.target === undefined && typeof raw.range === 'string') {
        raw.target = raw.range;
    }
    if (operation.op === 'renameSheet') {
        if (raw.oldName === undefined && raw.sheet !== undefined)
            raw.oldName = raw.sheet;
        if (raw.newName === undefined && raw.target !== undefined)
            raw.newName = raw.target;
    }
    if (['addSheet', 'deleteSheet', 'hideSheet', 'setTabColor', 'protectSheet', 'unprotectSheet', 'duplicateSheet'].includes(operation.op)) {
        if (raw.name === undefined && raw.sheet !== undefined)
            raw.name = raw.sheet;
        if (operation.op === 'duplicateSheet' && raw.newName === undefined && raw.target !== undefined)
            raw.newName = raw.target;
    }
    if (operation.op === 'filterToRange' && typeof raw.target === 'string' && !raw.target.includes('!') && !/^[A-Za-z]{1,3}\d+$/.test(raw.target)) {
        raw.target = `${raw.target}!A1`;
    }
    if (operation.op === 'crosstab') {
        // Flat metric fields -> metric object.
        if (raw.metric === undefined && (raw.metricColumn !== undefined || raw.metricFunction !== undefined)) {
            raw.metric = { column: raw.metricColumn, function: raw.metricFunction ?? 'sum' };
        }
    }
    if (operation.op === 'style' && raw.style && typeof raw.style === 'object') {
        // Models often emit exceljs-native alignment names instead of the DSL's.
        const style = raw.style;
        if (style.hAlign === undefined && style.horizontal !== undefined)
            style.hAlign = style.horizontal;
        if (style.vAlign === undefined && style.vertical !== undefined)
            style.vAlign = style.vertical;
    }
    if (['joinSheets', 'hideColumns'].includes(operation.op)) {
        for (const key of ['valueColumns', 'outputColumns', 'columns']) {
            const value = raw[key];
            if (value !== undefined && !Array.isArray(value)) {
                raw[key] = [value];
            }
        }
    }
    const out = {};
    for (const [key, value] of Object.entries(raw)) {
        if (key === 'cells' && value && typeof value === 'object') {
            out[key] = Array.isArray(value)
                ? value.map((id) => prefix(id))
                : Object.fromEntries(Object.entries(value).map(([id, content]) => [prefix(id), String(content)]));
        }
        else if (key === 'sheet') {
            out[key] = value ?? firstSheet;
        }
        else if (['range', 'source', 'target', 'start'].includes(key)) {
            out[key] = prefix(Array.isArray(value) ? value[0] : value);
        }
        else {
            out[key] = value;
        }
    }
    return out;
}
