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
/**
 * Evaluate every condition an assertion declares and require all of them.
 *
 * The earlier version returned as soon as one kind of condition matched, so
 * `{ expect, bold }` verified the value and silently ignored the style, and
 * `{ startsWith, fill }` ignored the fill. An assertion that names several
 * conditions is a conjunction — that is the whole point of naming them — and
 * checking only the first one produced exactly the false "achieved" this verifier
 * exists to prevent.
 */
function evaluateAssertion(assertion, cells, styleCells) {
    const normalized = normalizeCellId(assertion.id);
    const actual = cells[normalized] ?? cells[findKey(cells, normalized) ?? ''];
    const cell = styleCells?.get(normalized);
    const checks = [];
    if (assertion.expect !== undefined) {
        checks.push({
            label: t('值 期望 {expected} 实际 {actual}', {
                expected: formatValue(assertion.expect),
                actual: formatValue(actual),
            }),
            passed: assertion.expect === null
                ? actual === undefined || actual === ''
                : actual === assertion.expect,
        });
    }
    if (assertion.startsWith !== undefined) {
        checks.push({
            label: t('前缀 期望 {expected} 实际 {actual}', {
                expected: formatValue(assertion.startsWith),
                actual: formatValue(actual),
            }),
            passed: typeof actual === 'string' && actual.startsWith(assertion.startsWith),
        });
    }
    if (assertion.fill !== undefined) {
        checks.push({ label: t('填充色 期望 {value}', { value: assertion.fill }), passed: colorMatches(cell, assertion.fill) });
    }
    if (assertion.bold !== undefined) {
        checks.push({ label: t('加粗 期望 {value}', { value: assertion.bold }), passed: (cell?.font?.bold ?? false) === assertion.bold });
    }
    if (assertion.numberFormat !== undefined) {
        checks.push({ label: t('数字格式 期望 {value}', { value: assertion.numberFormat }), passed: cell?.numFmt === assertion.numberFormat });
    }
    if (assertion.wrapText !== undefined) {
        checks.push({ label: t('自动换行 期望 {value}', { value: assertion.wrapText }), passed: (cell?.alignment?.wrapText ?? false) === assertion.wrapText });
    }
    if (assertion.hAlign !== undefined) {
        checks.push({ label: t('水平对齐 期望 {value}', { value: assertion.hAlign }), passed: cell?.alignment?.horizontal === assertion.hAlign });
    }
    const failed = checks.filter((check) => !check.passed);
    const passed = checks.length > 0 && failed.length === 0;
    return {
        id: assertion.id,
        passed,
        detail: passed
            ? t('{id} 已满足全部 {count} 项要求', { id: assertion.id, count: checks.length })
            : t('{id} 未满足：{checks}', { id: assertion.id, checks: listJoin(failed.map((check) => check.label), 'semicolon') }),
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
