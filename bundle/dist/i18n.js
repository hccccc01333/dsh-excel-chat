const EN = {
    // health-report.ts
    '公式健康报告': 'Formula health report',
    '生成时间：{time}': 'Generated: {time}',
    '文件：{path}': 'File: {path}',
    '公式数：{count}': 'Formulas: {count}',
    '异常数：{count}': 'Anomalies: {count}',
    '健康分：{score}': 'Health score: {score}',
    '单元格': 'Cell',
    '类型': 'Kind',
    '说明': 'Detail',
    '健康分 {score}：{formulas} 个公式，{anomalies} 个异常，报告已写入 {path} 的「{sheet}」表': 'Health score {score}: {formulas} formulas, {anomalies} anomalies. The report was written to the "{sheet}" sheet of {path}.',
    // insight.ts
    '共 {sheets} 个工作表；{alerts} 个重点、{warns} 个提示。': '{sheets} worksheet(s): {alerts} alert(s), {warns} warning(s).',
    '{sheet}：{rows} 行数据，{columns} 列': '{sheet}: {rows} rows, {columns} columns',
    '，表头：{headers}': ', headers: {headers}',
    '，含 {count} 个公式': ', {count} formulas',
    '，有 {count} 处空值': ', {count} empty cells',
    '{sheet} 没有数据行，只有表头。': '{sheet} has no data rows, only a header.',
    '{label} 有 {count} 个空值（{percent}%）。': '{label} has {count} empty cells ({percent}%).',
    '{label} 值分布很集中，疑似存在大量重复值。': '{label} values are heavily concentrated — likely many duplicates.',
    '{label} 最大值 {max} 远高于均值 {mean}，疑似存在异常大值。': '{label} maximum {max} is far above the mean {mean} — likely an outlier.',
    '{label} 出现负数（最小值 {min}），请确认是否为退款/冲销。': '{label} contains negative values (minimum {min}); check whether these are refunds or reversals.',
    '{label} 存在首尾空格，建议 trimText。': '{label} has leading or trailing spaces; run trimText.',
    '{sheet} 含 {count} 个公式，可运行 excel_autofix 体检并修复。': '{sheet} contains {count} formulas; excel_autofix can check and repair them.',
    '有缺失值：用 excel_operate 的 fillMissing 补空值，或删除整空行。': 'Missing values: fill them with excel_operate fillMissing, or drop the empty rows.',
    '疑似重复：用 dedupeRows 按关键列去重。': 'Likely duplicates: dedupe with dedupeRows on the key column.',
    '发现异常/负值：建议先核对源数据，再用条件格式或图表突出展示。': 'Outliers or negatives: verify the source data first, then highlight them with conditional formatting or a chart.',
    '存在首尾空格：用 trimText 清理，再去做匹配/去重。': 'Leading or trailing spaces: clean them with trimText before matching or deduping.',
    '表里含公式：可运行 excel_autofix 体检并修复。': 'The workbook has formulas: excel_autofix can check and repair them.',
    '数据量较大：可用 excel_create_pivot / aggregateReport 做透视汇总，或 excel_create_chart 画图。': 'Large dataset: summarise it with excel_create_pivot / aggregateReport, or chart it with excel_create_chart.',
    '未发现明显数据问题；可继续做报表（report）、透视或图表。': 'No obvious data problems — you can move on to a report, a pivot table or a chart.',
    // menu.ts
    '未找到工作表。': 'No worksheet found.',
    '直接回复编号，或把示例话术发给我即可；做完不满意可以用 excel_undo 回滚。': 'Reply with a number, or just send one of the example prompts. If you do not like the result, excel_undo rolls it back.',
    '补空值': 'Fill blanks',
    '空值填充：固定值 / 向上取最近值 / 向左取最近值。': 'Fill blanks with a fixed value, from the cell above, or from the cell to the left.',
    '把 {sheet} 的{group}空值填 0': 'Fill the empty cells of {sheet} {group} with 0',
    '数据清洗': 'Clean data',
    '去重、补空值、删空行空列、去空格、大小写、分列。': 'Dedupe, fill blanks, drop empty rows and columns, trim, change case, split columns.',
    '把 {sheet} 按“{group}”去重，名称去掉首尾空格': 'Dedupe {sheet} by "{group}" and trim the names',
    '公式体检 + 自愈': 'Formula check + self-healing',
    '检查公式有没有被弄坏，不对的自动修复并复验。': 'Check whether the formulas were broken, repair the wrong ones and re-validate.',
    '检查 {sheet} 的公式有没有错，不对的修掉': 'Check the formulas in {sheet} and fix the wrong ones',
    '一键经营报表': 'One-shot business report',
    '排序 + 分类汇总 + 动态透视 + 筛选 + 样式 + 冻结一步完成。': 'Sort, subtotal, dynamic summary, filter, styling and freeze panes in one step.',
    '用 report 给 {sheet} 做经营报表：按“{group}”分组，“{metric}”合计': 'Build a business report for {sheet} with report: group by "{group}", total "{metric}"',
    '动态透视汇总': 'Dynamic summary',
    '按分组字段生成实时 SUMIFS 联动汇总表。': 'Generate a live SUMIFS summary table grouped by a column.',
    '按“{group}”汇总“{metric}”，输出到新表': 'Summarise "{metric}" by "{group}" into a new sheet',
    '原生透视表': 'Native pivot table',
    'Excel 原生数据透视表，可交互、可刷新。': 'A native Excel pivot table — interactive and refreshable.',
    '给 {sheet} 建透视表：行字段 {row}，值 {col} 求和': 'Build a pivot table for {sheet}: rows {row}, sum of {col}',
    '图表': 'Chart',
    '柱状图/折线图/饼图，可改标题、图例、坐标轴。': 'Bar, line or pie charts; titles, legends and axes can be edited.',
    '给 {sheet} 生成柱状图：“{group}”为分类，“{metric}”为数值': 'Chart {sheet}: "{group}" as categories, "{metric}" as values',
    '批量通知': 'Bulk notifications',
    '用占位符模板给每一行生成一条通知。': 'Generate one message per row from a placeholder template.',
    '用“通知模板”表给 {sheet} 每行生成一条发货通知': 'Use the template sheet to generate one notice per row of {sheet}',
    '岗位模板': 'Role templates',
    '运营 / 产品 / 数分三种报表模板，按岗位一键套用。': 'Operations / product / data-analysis report templates, applied in one step.',
    '我是运营，帮我把 {sheet} 做成运营报表': 'I work in operations — turn {sheet} into an operations report',
    // autofix.ts
    '体检：修复前 {before} 个异常，修复后 {after} 个。': 'Health check: {before} anomalies before repair, {after} after.',
    '已修复 {count} 处：{ids}': 'Repaired {count}: {ids}',
    '未自动修复 {count} 处；可提供表格结构后启用 LLM 修复，或人工核对。': '{count} left unrepaired; supply the table structure to enable LLM repair, or check them by hand.',
    '未发现公式异常，无需修复。': 'No formula anomalies found — nothing to repair.',
    '输出文件：{path}': 'Output file: {path}',
    // doctor.ts
    'node {version}（要求 ^22.19 || >=24）': 'node {version} (requires ^22.19 || >=24)',
    '未发现 dsh profile，跳过宿主包隔离检查': 'No dsh profile found; skipping the host-package isolation check',
    '{dir}：宿主包未出现在 dependencies（符合隔离要求）': '{dir}: host packages are not in dependencies (isolation is intact)',
    '{dir}：宿主包被安装为 dependencies：{list}（会导致所有工具调用失败）': '{dir}: host packages are installed as dependencies: {list} (this breaks every tool call)',
    '临时工作簿读取+公式体检正常（公式 {formulas} 个，异常 {anomalies} 个）': 'Temp workbook read + formula check OK ({formulas} formulas, {anomalies} anomalies)',
    '引擎冒烟失败：{error}': 'Engine smoke test failed: {error}',
    // operations.ts (messages only — labels written into the workbook stay as they are)
    'freezeFormulas 转换 {frozen} 个公式，跳过 {skipped} 个无缓存结果的（先在 Excel 中打开计算后可再转）': 'freezeFormulas converted {frozen} formulas and skipped {skipped} with no cached result (open the file in Excel to calculate, then convert)',
    'freezeFormulas 转换 {frozen} 个公式为缓存值': 'freezeFormulas converted {frozen} formula(s) to their cached values',
    'copyRange valuesOnly：部分公式无缓存结果，已按公式复制': 'copyRange valuesOnly: some formulas had no cached result and were copied as formulas',
    // index.ts — slash commands (the tool descriptions and the system-prompt
    // section are model-facing and stay as they are)
    '就地修改本地 Excel 文件的一个单元格（自动备份 .bak + 审计日志 + 公式体检）': 'Edit one cell of a local .xlsx in place (auto .bak backup + audit log + formula check)',
    '需要 {"path","cell","value"}': 'requires {"path","cell","value"}',
    '已就地保存 {cell} = {value}（备份 {backup}，公式异常 {anomalies}）': 'Saved {cell} = {value} in place (backup {backup}, {anomalies} formula anomalies)',
    '回滚本地 Excel 文件最近一次就地修改': 'Roll back the most recent in-place edit of a local .xlsx',
    '需要 {"path"}': 'requires {"path"}',
    '安装自检：检查宿主包隔离、Node 版本与 Excel 引擎冒烟': 'Install self-check: host-package isolation, Node version and an Excel engine smoke test',
    '可选：{"profile":"D:\\\\profile-dir"}': 'optional: {"profile":"D:\\\\profile-dir"}',
    // live-edit.ts
    '没有可回滚的编辑记录': 'No edit record to roll back',
    '已回滚 {count} 处修改': 'Rolled back {count} change(s)',
    // llm-benchmark.ts
    '{op}.{key}：期望 {expected}，实际 {actual}': '{op}.{key}: expected {expected}, got {actual}',
    // preview.ts
    '表 {sheet}：展示 {rows} 行 × {cols} 列': 'Sheet {sheet}: {rows} rows × {cols} columns',
    '（范围 {range}）': ' (range {range})',
    '，HTML 预览：{path}': ', HTML preview: {path}',
    '{sheet} 预览': '{sheet} preview',
    // verifier.ts
    '确定性断言全部通过（{passed}/{total}）': 'All deterministic assertions passed ({passed}/{total})',
    '没有可执行的确定性断言': 'No deterministic assertions to run',
    '确定性断言未全部通过（{passed}/{total}）：{failures}': 'Deterministic assertions failed ({passed}/{total}): {failures}',
    '{id} 已满足期望值': '{id} matches the expected value',
    '{id} 期望 {expected}，实际 {actual}': '{id} expected {expected}, got {actual}',
    '{id} 已满足前缀要求': '{id} matches the expected prefix',
    '{id} 期望以 {expected} 开头，实际 {actual}': '{id} expected to start with {expected}, got {actual}',
    '填充色={value}': 'fill={value}',
    '加粗={value}': 'bold={value}',
    '数字格式={value}': 'numberFormat={value}',
    '自动换行={value}': 'wrapText={value}',
    '水平对齐={value}': 'hAlign={value}',
    '{id} 已满足样式要求': '{id} matches the style requirements',
    '{id} {detail}': '{id} {detail}',
    '样式不符合：{checks}': 'style mismatch: {checks}',
    '不存在或没有可检查的样式': 'is missing or has no checkable style',
    '缺失': 'missing',
    // semantic.ts (the role-name heuristics in this module are patterns, not messages)
    '未识别': 'unrecognised',
    '无': 'none',
    '{sheet}：粒度={grain}': '{sheet}: grain={grain}',
    '时间={list}': 'time={list}',
    '维度={list}': 'dimension={list}',
    '指标={list}': 'measure={list}',
    '标识={list}': 'id={list}',
    '派生={list}': 'derived={list}',
    '可关联：{keys}': 'joinable: {keys}',
    '未发现跨表关联键': 'no cross-sheet join keys found',
    // menu.ts (fallback label when a column has no header)
    '{column} 列': 'column {column}',
    // agent.ts
    '{count} 个公式异常': '{count} formula anomalies',
    '计划无效：{message}。请修正后重新规划。': 'Invalid plan: {message}. Fix it and plan again.',
    '{message}（第 {round} 轮计划：{ops}）': '{message} (round {round} plan: {ops})',
    '执行出错：{message}。请修正计划后重新规划。': 'Execution failed: {message}. Fix the plan and plan again.',
    '公式无异常': 'no formula anomalies',
    '仍有 {count} 个公式异常': 'still {count} formula anomalies',
    '{anomalies}；文件{changed}实质变化': '{anomalies}; the file {changed} change substantially',
    '有': 'did',
    '没有': 'did not',
    '{reason}（确定性校验：{note}）': '{reason} (deterministic check: {note})',
    '{reason}（规划器断言未过 {passed}/{total}：{failures}）': '{reason} (planner assertions failed {passed}/{total}: {failures})',
    '{sheet}：{rows} 行 × {columns} 列': '{sheet}: {rows} rows × {columns} columns',
    '，表头 {headers}': ', headers {headers}',
    // plan-schema.ts — salvage notes
    '断言不是数组，已丢弃': 'assertions is not an array; dropped',
    '断言[{index}] 不是对象，已丢弃': 'assertion[{index}] is not an object; dropped',
    '断言[{index}] 缺少 id，已丢弃': 'assertion[{index}] has no id; dropped',
    '断言[{index}] 的 id 已补工作表前缀': 'assertion[{index}] id was given a sheet prefix',
    '断言[{index}] 的 expect 类型不支持，已忽略该字段': 'assertion[{index}] expect has an unsupported type; field ignored',
    '断言[{index}] 的 startsWith 必须是非空字符串，已忽略': 'assertion[{index}] startsWith must be a non-empty string; ignored',
    '断言[{index}] 没有可检查字段，已丢弃': 'assertion[{index}] has nothing to check; dropped',
    '{op} 的 {key} 已补工作表前缀': '{op} {key} was given a sheet prefix',
    '{op} 已补默认工作表': '{op} was given the default sheet',
    '{op} 的 sheet 已匹配为 {matched}': '{op} sheet was matched to {matched}',
    '{op} 的 name 已匹配为 {matched}': '{op} name was matched to {matched}',
    'renameSheet 的 oldName 已匹配为 {matched}': 'renameSheet oldName was matched to {matched}',
    '{op} 的 {key} 已转为字符串': '{op} {key} was coerced to a string',
    '{op} 的 {field}[{index}].column 已转为字符串': '{op} {field}[{index}].column was coerced to a string',
    'freezePanes 已从 range {range} 推导 row/column': 'freezePanes derived row/column from range {range}',
    'crosstab 的 metricColumn/metricFunction 已合并为 metric 对象': 'crosstab metricColumn/metricFunction were merged into a metric object',
    'crosstab 的 metric.function 已补默认 sum': 'crosstab metric.function defaulted to sum',
    'crosstab 的 metric.column 已转为字符串': 'crosstab metric.column was coerced to a string',
    'style 的 horizontal 已改名为 hAlign': 'style horizontal was renamed to hAlign',
    'style 的 vertical 已改名为 vAlign': 'style vertical was renamed to vAlign',
    '{op} 的 target 已扩展为 {target}': '{op} target was expanded to {target}',
    '{op} 的 {field} 已包装为数组': '{op} {field} was wrapped in an array',
    'fillMissing 已补 mode=value': 'fillMissing defaulted to mode=value',
    'fillMissing 的 fillValue 已改为 value': 'fillMissing fillValue was renamed to value',
    'filterToRange 的 target 已补 !A1': 'filterToRange target was given !A1',
    // plan-schema.ts — validation failures (kind is carried by PlanSchemaError)
    '第 {step} 步没有 operations 数组': 'step {step} has no operations array',
    '第 {step} 步第 {op} 个操作缺少 op 字段': 'step {step} operation {op} has no op field',
    '{op} 的 {field}[{index}] 必须是对象': '{op} {field}[{index}] must be an object',
    '{op} 的 {field}[{index}].column 缺失': '{op} {field}[{index}].column is missing',
    'crosstab 的 metric.function 不支持：{value}': 'crosstab metric.function is not supported: {value}',
    '{op} 缺少必填数字 {key}': '{op} is missing the required number {key}',
    '{op} 的 {key} 必须是数字': '{op} {key} must be a number',
    '{op} 缺少必填数组 {field}': '{op} is missing the required array {field}',
    '{op} 缺少必填字段 {field}': '{op} is missing the required field {field}',
    // failure-taxonomy.ts
    '规划器/计划结构错误：{message}': 'Planner / plan-structure error: {message}',
    '参数错误：{message}': 'Argument error: {message}',
    '执行异常：{message}': 'Execution error: {message}',
    '验证器判定目标已达成，但断言只过 {passed}/{total}，完整性异常 {integrity}': 'The verifier claimed success, but only {passed}/{total} checks passed and integrity found {integrity} anomalies',
    '{rounds} 轮重规划后仍未达成目标，第一轮失败后没有纠正': 'Still not achieved after {rounds} replanning rounds; the first failure was never corrected',
    '没有执行任何操作': 'No operation was executed',
    '期望操作都已执行且参数一致，但断言未过，可能是列/表/指标语义理解偏差': 'The expected operations ran with matching arguments but the checks failed — likely a column/table/measure misunderstanding',
    '缺少 {missing}，改用通用操作 {unexpected}': 'Missing {missing}; fell back to generic operations {unexpected}',
    '缺关键步骤：{missing}': 'Missing key steps: {missing}',
    '期望操作 {expected}，实际执行 {actual}': 'Expected {expected}, actually executed {actual}',
    // explain.ts — function descriptions, translated at lookup time
    '求和': 'sum',
    '求平均': 'average',
    '计数（只数数字）': 'count (numbers only)',
    '计数（非空）': 'count (non-empty)',
    '取最大值': 'maximum',
    '取最小值': 'minimum',
    '取中位数': 'median',
    '求乘积': 'product',
    '按条件求和': 'sum with one condition',
    '按多个条件求和': 'sum with multiple conditions',
    '按条件计数': 'count with one condition',
    '按多个条件计数': 'count with multiple conditions',
    '按条件求平均': 'average with one condition',
    '按多个条件求平均': 'average with multiple conditions',
    '分类汇总': 'subtotal',
    '纵向查找（按首列找并返回指定列）': 'vertical lookup (match the first column, return another column)',
    '查找并返回匹配值': 'lookup and return the matching value',
    '按行列位置取值': 'value at a row/column position',
    '查找目标所在位置': 'position of a value',
    '条件判断，成立返回一个值、否则返回另一个值': 'conditional: one value when true, another when false',
    '出错时返回替代值': 'fallback value when the expression errors',
    '查不到时返回替代值': 'fallback value when a lookup finds nothing',
    '返回当天日期': "today's date",
    '返回当前日期时间': 'current date and time',
    '取年份': 'year',
    '取月份': 'month',
    '取日': 'day of month',
    '按年月日拼日期': 'build a date from year/month/day',
    '计算两个日期的间隔': 'interval between two dates',
    '返回某月最后一天': 'last day of a month',
    '按格式转文本': 'format a value as text',
    '四舍五入': 'round',
    '向上取整': 'round up',
    '向下取整': 'round down',
    '取整': 'integer part',
    '取余数': 'remainder',
    '取绝对值': 'absolute value',
    '去掉多余空格': 'strip extra spaces',
    '计算字符数': 'character count',
    '取左侧若干个字符': 'leftmost characters',
    '取右侧若干个字符': 'rightmost characters',
    '从中间取字符': 'characters from the middle',
    '拼接文本': 'concatenate text',
    '替换文本': 'replace text',
    '使用了函数：{list}': 'Functions used: {list}',
    '引用区域：{list}': 'Referenced ranges: {list}',
    '涉及跨表引用：{list}': 'Cross-sheet references: {list}',
    '包含算术运算（加/减/乘/除/乘方）': 'Contains arithmetic (+ − × ÷ power)',
    '包含比较判断': 'Contains a comparison',
    '这是一个 {functions} 公式：{descriptions}。': 'This is a {functions} formula: {descriptions}.',
    '{name}（{description}）': '{name} ({description})',
    '这是一个引用其他单元格/区域参与计算或比较的公式。': 'This formula references other cells or ranges in a calculation or comparison.',
    '这是一个常量或简单表达式。': 'This is a constant or a simple expression.',
};
let language = 'zh';
/** Set the language for user-facing messages. Called once from the plugin config. */
export function setLanguage(next) {
    language = next;
}
export function getLanguage() {
    return language;
}
/**
 * Translate a user-facing message, substituting `{name}` placeholders. Falls
 * back to the Chinese source when no translation is registered, so a missing
 * entry shows the original wording rather than a key.
 */
export function t(zh, params) {
    const template = language === 'en' ? (EN[zh] ?? zh) : zh;
    if (params === undefined)
        return template;
    return template.replace(/\{(\w+)\}/g, (match, key) => key in params ? String(params[key]) : match);
}
/** True when an English translation is registered for this source text. */
export function hasTranslation(zh) {
    return zh in EN;
}
/**
 * The list separator itself, for callers that join an array directly. The
 * separator is language-specific — Chinese uses the enumeration comma `、` and
 * the full-width semicolon `；`, English the plain comma and semicolon — so
 * hard-coding either one leaves English sentences punctuated like Chinese
 * (`A、B、C`).
 */
export function listSeparator(style = 'comma') {
    if (language === 'en')
        return style === 'semicolon' ? '; ' : ', ';
    return style === 'semicolon' ? '；' : '、';
}
/** Join a list for display with the language's own separator. */
export function listJoin(items, style = 'comma') {
    return items.join(listSeparator(style));
}
/** Source strings that have no English translation yet, for the coverage test. */
export function untranslated(keys) {
    return keys.filter((key) => !(key in EN));
}
