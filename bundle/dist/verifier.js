import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { t, listJoin } from './i18n.js';
import { normalizeCellId } from './score.js';
import { readWorkbookCells, stripPivotTableParts } from './workbook.js';
/**
 * Evaluate workbook assertions without an LLM.
 *
 * Cell-content assertions use the normalized workbook cell map. Presentation
 * assertions load the workbook once and require every declared presentation
 * property on a check to match.
 */
export async function verifyWorkbookAssertions(path, assertions) {
    const cells = await readWorkbookCells(await readFile(path));
    const needsStyles = assertions.some((assertion) => hasStyleAssertion(assertion));
    const styleCells = needsStyles ? await loadStyleCells(path) : null;
    const results = assertions.map((assertion) => evaluateAssertion(assertion, cells, styleCells));
    const failures = results.filter((result) => !result.passed).map((result) => result.detail);
    const passed = results.length - failures.length;
    const achieved = results.length > 0 && failures.length === 0;
    const reason = achieved
        ? t('确定性断言全部通过（{passed}/{total}）', { passed, total: results.length })
        : results.length === 0
            ? t('没有可执行的确定性断言')
            : t('确定性断言未全部通过（{passed}/{total}）：{failures}', {
                passed,
                total: results.length,
                failures: listJoin(failures.slice(0, 4), 'semicolon'),
            });
    return { achieved, passed, total: results.length, failures, assertions: results, reason };
}
function evaluateAssertion(assertion, cells, styleCells) {
    const normalized = normalizeCellId(assertion.id);
    const actual = cells[normalized] ?? cells[findKey(cells, normalized) ?? ''];
    if (assertion.expect !== undefined) {
        const matches = assertion.expect === null
            ? actual === undefined || actual === ''
            : actual === assertion.expect;
        return {
            id: assertion.id,
            passed: matches,
            detail: matches
                ? t('{id} 已满足期望值', { id: assertion.id })
                : t('{id} 期望 {expected}，实际 {actual}', {
                    id: assertion.id,
                    expected: formatValue(assertion.expect),
                    actual: formatValue(actual),
                }),
        };
    }
    if (assertion.startsWith !== undefined) {
        const matches = typeof actual === 'string' && actual.startsWith(assertion.startsWith);
        return {
            id: assertion.id,
            passed: matches,
            detail: matches
                ? t('{id} 已满足前缀要求', { id: assertion.id })
                : t('{id} 期望以 {expected} 开头，实际 {actual}', {
                    id: assertion.id,
                    expected: formatValue(assertion.startsWith),
                    actual: formatValue(actual),
                }),
        };
    }
    const cell = styleCells?.get(normalized);
    const checks = [
        ...(assertion.fill !== undefined ? [{ label: t('填充色={value}', { value: assertion.fill }), passed: colorMatches(cell, assertion.fill) }] : []),
        ...(assertion.bold !== undefined ? [{ label: t('加粗={value}', { value: assertion.bold }), passed: (cell?.font?.bold ?? false) === assertion.bold }] : []),
        ...(assertion.numberFormat !== undefined ? [{ label: t('数字格式={value}', { value: assertion.numberFormat }), passed: cell?.numFmt === assertion.numberFormat }] : []),
        ...(assertion.wrapText !== undefined ? [{ label: t('自动换行={value}', { value: assertion.wrapText }), passed: (cell?.alignment?.wrapText ?? false) === assertion.wrapText }] : []),
        ...(assertion.hAlign !== undefined ? [{ label: t('水平对齐={value}', { value: assertion.hAlign }), passed: cell?.alignment?.horizontal === assertion.hAlign }] : []),
    ];
    const passed = checks.length > 0 && checks.every((check) => check.passed);
    const failedChecks = checks.filter((check) => !check.passed).map((check) => check.label);
    return {
        id: assertion.id,
        passed,
        detail: passed
            ? t('{id} 已满足样式要求', { id: assertion.id })
            : t('{id} {detail}', {
                id: assertion.id,
                detail: cell
                    ? t('样式不符合：{checks}', { checks: listJoin(failedChecks) })
                    : t('不存在或没有可检查的样式'),
            }),
    };
}
function hasStyleAssertion(assertion) {
    return assertion.fill !== undefined
        || assertion.bold !== undefined
        || assertion.numberFormat !== undefined
        || assertion.wrapText !== undefined
        || assertion.hAlign !== undefined;
}
function colorMatches(cell, expected) {
    const actual = cell?.fill?.type === 'pattern'
        ? cell.fill.fgColor?.argb
        : undefined;
    return actual !== undefined && actual.toUpperCase().endsWith(expected.toUpperCase());
}
function formatValue(value) {
    return value === undefined ? t('缺失') : JSON.stringify(value);
}
async function loadStyleCells(path) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(stripPivotTableParts(await readFile(path)));
    const cells = new Map();
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                cells.set(normalizeCellId(`${sheet.name}!${cell.address}`), cell);
            });
        });
    });
    return cells;
}
function findKey(cells, normalized) {
    return Object.keys(cells).find((key) => normalizeCellId(key) === normalized);
}
