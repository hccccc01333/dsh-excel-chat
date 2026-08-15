import { profileWorkbook } from './profile.js';
/**
 * Turn a workbook into a plain-language capability menu: a one-line digest of
 * what the file contains plus concrete next steps, each with an example prompt
 * the user can just confirm. This is the "let the agent speak first" entry
 * point for users who cannot describe what they want.
 */
export async function buildWorkbookMenu(path, sheet) {
    const profile = await profileWorkbook(path, sheet);
    const primary = profile.sheets[0];
    if (!primary) {
        return { summary: '未找到工作表。', suggestions: [], note: '' };
    }
    return {
        summary: summarize(profile, primary),
        suggestions: buildSuggestions(profile, primary),
        note: '直接回复编号，或把示例话术发给我即可；做完不满意可以用 excel_undo 回滚。',
    };
}
function summarize(profile, primary) {
    const parts = [];
    for (const sheet of profile.sheets) {
        const headers = sheet.columns.filter((column) => column.header).map((column) => column.header);
        let line = `${sheet.sheet}：${sheet.dataRows} 行数据，${sheet.columnCount} 列`;
        if (headers.length > 0)
            line += `，表头：${headers.slice(0, 6).join(' / ')}${headers.length > 6 ? ' …' : ''}`;
        if (sheet.formulaCells > 0)
            line += `，含 ${sheet.formulaCells} 个公式`;
        const missing = sheet.columns.reduce((sum, column) => sum + column.missing, 0);
        if (missing > 0)
            line += `，有 ${missing} 处空值`;
        parts.push(line);
    }
    return parts.join('\n');
}
function buildSuggestions(profile, primary) {
    const sheet = primary.sheet;
    const groupColumn = primary.columns.find((column) => column.dtype === 'string' && column.header) ?? primary.columns[0];
    const metricColumn = primary.columns.find((column) => column.dtype === 'number') ?? groupColumn;
    const groupHeader = groupColumn?.header ?? `${groupColumn?.column ?? 'A'} 列`;
    const metricHeader = metricColumn?.header ?? `${metricColumn?.column ?? 'B'} 列`;
    const hasMissing = primary.columns.some((column) => column.missing > 0);
    const hasFormulas = primary.formulaCells > 0;
    const suggestions = [];
    if (hasMissing) {
        suggestions.push({
            id: 'fillMissing',
            title: '补空值',
            description: '空值填充：固定值 / 向上取最近值 / 向左取最近值。',
            example: `把 ${sheet} 的${groupHeader}空值填 0`,
        });
    }
    suggestions.push({
        id: 'clean',
        title: '数据清洗',
        description: '去重、补空值、删空行空列、去空格、大小写、分列。',
        example: `把 ${sheet} 按“${groupHeader}”去重，名称去掉首尾空格`,
    });
    if (hasFormulas) {
        suggestions.push({
            id: 'health',
            title: '公式体检 + 自愈',
            description: '检查公式有没有被弄坏，不对的自动修复并复验。',
            example: `检查 ${sheet} 的公式有没有错，不对的修掉`,
        });
    }
    suggestions.push({
        id: 'report',
        title: '一键经营报表',
        description: '排序 + 分类汇总 + 动态透视 + 筛选 + 样式 + 冻结一步完成。',
        example: `用 report 给 ${sheet} 做经营报表：按“${groupHeader}”分组，“${metricHeader}”合计`,
    }, {
        id: 'aggregate',
        title: '动态透视汇总',
        description: '按分组字段生成实时 SUMIFS 联动汇总表。',
        example: `按“${groupHeader}”汇总“${metricHeader}”，输出到新表`,
    }, {
        id: 'pivot',
        title: '原生透视表',
        description: 'Excel 原生数据透视表，可交互、可刷新。',
        example: `给 ${sheet} 建透视表：行字段 ${groupColumn?.column ?? 'A'}，值 ${metricColumn?.column ?? 'B'} 求和`,
    }, {
        id: 'chart',
        title: '图表',
        description: '柱状图/折线图/饼图，可改标题、图例、坐标轴。',
        example: `给 ${sheet} 生成柱状图：“${groupHeader}”为分类，“${metricHeader}”为数值`,
    }, {
        id: 'mail',
        title: '批量通知',
        description: '用占位符模板给每一行生成一条通知。',
        example: `用“通知模板”表给 ${sheet} 每行生成一条发货通知`,
    }, {
        id: 'preset',
        title: '岗位模板',
        description: '运营 / 产品 / 数分三种报表模板，按岗位一键套用。',
        example: `我是运营，帮我把 ${sheet} 做成运营报表`,
    });
    return suggestions;
}
