import { t } from './i18n.js';
import { profileWorkbook } from './profile.js';
import { analyzeWorkbook } from './analysis.js';
/**
 * Heuristic data insight report (ExcelGenius2-style "upload -> summary +
 * anomalies"): per-sheet one-liner, missing/duplicate/outlier/normalization
 * findings, and concrete next-step suggestions. Deterministic, no LLM needed.
 *
 * Two kinds of finding come back. Most answer "is this data trustworthy" — missing
 * values, duplicates, outliers. The rest come from `analysis.ts` and answer "what do the
 * numbers say" — a trend across periods, a measure concentrated in a few categories.
 * The second kind was missing entirely: a sheet with a clean 100→260 ramp and one region
 * four times the size of the others produced no findings at all.
 */
export async function buildWorkbookInsight(path, sheet) {
    const profile = await profileWorkbook(path, sheet);
    const sheetInsights = profile.sheets.map((entry) => insightForSheet(entry));
    const analysis = await analyzeWorkbook(path, sheet);
    for (const finding of analysis) {
        const target = sheetInsights.find((entry) => entry.sheet === finding.sheet);
        if (target)
            target.findings.push(analysisToInsight(finding));
    }
    const findings = sheetInsights.flatMap((entry) => entry.findings);
    const suggestions = buildSuggestions(profile, findings);
    return {
        summary: t('共 {sheets} 个工作表；{alerts} 个重点、{warns} 个提示。', {
            sheets: profile.sheetCount,
            alerts: findings.filter((f) => f.severity === 'alert').length,
            warns: findings.filter((f) => f.severity === 'warn').length,
        }),
        sheets: sheetInsights,
        suggestions,
    };
}
/**
 * An analytical finding, presented as an insight.
 *
 * `info`, not `warn`: a trend or a concentrated measure is not a problem to fix. Marking
 * it as one would put "your sales grew" in the same list as "you have 40% missing
 * values", and the alerts count at the top would stop meaning anything.
 */
function analysisToInsight(finding) {
    return { severity: 'info', category: finding.kind, message: finding.message };
}
function insightForSheet(sheet) {
    const findings = [];
    const headers = sheet.columns.filter((column) => column.header).map((column) => column.header);
    let summary = t('{sheet}：{rows} 行数据，{columns} 列', {
        sheet: sheet.sheet,
        rows: sheet.dataRows,
        columns: sheet.columnCount,
    });
    if (headers.length > 0) {
        summary += t('，表头：{headers}', {
            headers: `${headers.slice(0, 6).join(' / ')}${headers.length > 6 ? ' …' : ''}`,
        });
    }
    if (sheet.formulaCells > 0)
        summary += t('，含 {count} 个公式', { count: sheet.formulaCells });
    if (sheet.dataRows === 0) {
        findings.push({ severity: 'info', category: 'empty', message: t('{sheet} 没有数据行，只有表头。', { sheet: sheet.sheet }) });
    }
    for (const column of sheet.columns) {
        if (!column.header)
            continue;
        const label = `${sheet.sheet}!${column.column}（${column.header}）`;
        if (column.missing > 0 && column.nonEmpty + column.missing > 0) {
            const ratio = column.missing / (column.nonEmpty + column.missing);
            findings.push({
                severity: ratio >= 0.2 ? 'warn' : 'info',
                category: 'missing',
                message: t('{label} 有 {count} 个空值（{percent}%）。', {
                    label,
                    count: column.missing,
                    percent: Math.round(ratio * 100),
                }),
            });
        }
        if (column.nonEmpty > 3 && column.uniqueCapped) {
            findings.push({
                severity: 'warn',
                category: 'duplicate',
                message: t('{label} 值分布很集中，疑似存在大量重复值。', { label }),
            });
        }
        if (column.dtype === 'number' && column.mean !== undefined && column.max !== undefined && column.min !== undefined && column.mean > 0) {
            if (column.max > column.mean * 5 && column.max - column.mean > column.mean) {
                findings.push({
                    severity: 'warn',
                    category: 'outlier',
                    message: t('{label} 最大值 {max} 远高于均值 {mean}，疑似存在异常大值。', {
                        label,
                        max: column.max,
                        mean: column.mean,
                    }),
                });
            }
            if (column.min < 0 && /(金额|amount|price|cost|revenue|sales|profit|总额|费用)/i.test(column.header)) {
                findings.push({
                    severity: 'warn',
                    category: 'negative',
                    message: t('{label} 出现负数（最小值 {min}），请确认是否为退款/冲销。', { label, min: column.min }),
                });
            }
        }
        if (column.dtype === 'string' && column.samples.some((sample) => sample !== sample.trim())) {
            findings.push({
                severity: 'info',
                category: 'whitespace',
                message: t('{label} 存在首尾空格，建议 trimText。', { label }),
            });
        }
    }
    if (sheet.formulaCells > 0) {
        findings.push({
            severity: 'info',
            category: 'formula',
            message: t('{sheet} 含 {count} 个公式，可运行 excel_autofix 体检并修复。', {
                sheet: sheet.sheet,
                count: sheet.formulaCells,
            }),
        });
    }
    return { sheet: sheet.sheet, summary, findings };
}
function buildSuggestions(profile, findings) {
    const suggestions = [];
    if (findings.some((f) => f.category === 'missing'))
        suggestions.push(t('有缺失值：用 excel_operate 的 fillMissing 补空值，或删除整空行。'));
    if (findings.some((f) => f.category === 'duplicate'))
        suggestions.push(t('疑似重复：用 dedupeRows 按关键列去重。'));
    if (findings.some((f) => f.category === 'outlier' || f.category === 'negative'))
        suggestions.push(t('发现异常/负值：建议先核对源数据，再用条件格式或图表突出展示。'));
    if (findings.some((f) => f.category === 'whitespace'))
        suggestions.push(t('存在首尾空格：用 trimText 清理，再去做匹配/去重。'));
    if (findings.some((f) => f.category === 'trend'))
        suggestions.push(t('发现趋势：可用 excel_create_chart 画折线图，或 report 做汇总表。'));
    if (findings.some((f) => f.category === 'concentration'))
        suggestions.push(t('发现集中度：可用 crosstab 做二维对比，或 highlightRows 把头部项标出来。'));
    if (profile.sheets.some((s) => s.formulaCells > 0))
        suggestions.push(t('表里含公式：可运行 excel_autofix 体检并修复。'));
    if (profile.sheets.some((s) => s.dataRows > 20))
        suggestions.push(t('数据量较大：可用 excel_create_pivot / aggregateReport 做透视汇总，或 excel_create_chart 画图。'));
    if (suggestions.length === 0)
        suggestions.push(t('未发现明显数据问题；可继续做报表（report）、透视或图表。'));
    return suggestions;
}
