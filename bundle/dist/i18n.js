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
/** Source strings that have no English translation yet, for the coverage test. */
export function untranslated(keys) {
    return keys.filter((key) => !(key in EN));
}
