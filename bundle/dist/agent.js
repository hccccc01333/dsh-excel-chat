import { copyFile, mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { t, listJoin, listSeparator } from './i18n.js';
import { PlanSchemaError, sanitizeAssertions, sanitizePlan } from './plan-schema.js';
import { profileWorkbook } from './profile.js';
import { buildWorkbookSemanticProfile } from './semantic.js';
import { runExcelTask } from './task.js';
import { verifyWorkbookAssertions } from './verifier.js';
import { readWorkbookCells, validateWorkbookFile, workbookFingerprint } from './workbook.js';
/**
 * Goal-driven agent loop (Plan -> Act -> Observe -> Verify -> Replan):
 * the planner proposes operation steps for the goal, `runExcelTask` executes
 * them with per-step formula verification and deterministic repair, an LLM
 * verifier checks whether the goal is achieved, and the loop replans up to
 * maxRounds times when it is not.
 */
export async function runAgentTask(path, options) {
    const maxRounds = options.maxRounds ?? 2;
    if (maxRounds < 1)
        throw new Error('maxRounds must be at least 1');
    const dir = await mkdtemp(join(tmpdir(), 'vera-agent-'));
    let currentPath = path;
    const rounds = [];
    let achieved = false;
    let previousPlan;
    let previousResult;
    let verifierNote;
    for (let round = 1; round <= maxRounds; round++) {
        const beforeProfile = await profileWorkbook(currentPath);
        const semanticProfile = await buildWorkbookSemanticProfile(currentPath);
        const beforeValidation = await validateWorkbookFile(currentPath);
        const beforeFingerprint = await workbookFingerprint(currentPath);
        const planContext = {
            goal: options.goal,
            path: currentPath,
            round,
            sheetNames: beforeProfile.sheets.map((sheet) => sheet.sheet),
            profileSummary: summarizeProfile(beforeProfile),
            semanticSummary: semanticProfile.summary,
            validationSummary: t('{count} 个公式异常', { count: beforeValidation.anomalies.length }),
            previousPlan,
            previousResult,
            verifierNote,
        };
        let plan;
        let planAssertions = [];
        try {
            const rawPlan = await options.planner.plan(planContext);
            const steps = Array.isArray(rawPlan) ? rawPlan : rawPlan.steps;
            if (!steps || steps.length === 0)
                throw new Error('planner returned an empty plan');
            plan = sanitizePlan(steps, planContext.sheetNames).steps;
            if (!Array.isArray(rawPlan) && rawPlan.assertions !== undefined) {
                const sanitized = sanitizeAssertions(rawPlan.assertions, planContext.sheetNames);
                planAssertions = sanitized.assertions;
            }
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            if (round < maxRounds) {
                verifierNote = t('计划无效：{message}。请修正后重新规划。', { message });
                previousPlan = undefined;
                previousResult = undefined;
                continue;
            }
            throw wrapRoundError(error, t('{message}（第 {round} 轮计划：{ops}）', { message, round, ops: summarizePlanOps(plan ?? []) }));
        }
        const roundOut = join(dir, `round-${round}.xlsx`);
        let result;
        try {
            result = await runExcelTask(currentPath, plan, roundOut);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            if (round < maxRounds) {
                verifierNote = t('执行出错：{message}。请修正计划后重新规划。', { message });
                previousPlan = plan;
                previousResult = undefined;
                continue;
            }
            throw wrapRoundError(error, t('{message}（第 {round} 轮计划：{ops}）', { message, round, ops: summarizePlanOps(plan) }));
        }
        const afterProfile = await profileWorkbook(result.outputPath);
        const afterValidation = await validateWorkbookFile(result.outputPath);
        const cellSnapshot = await cellSnapshotOf(result.outputPath);
        const changed = (await workbookFingerprint(result.outputPath)) !== beforeFingerprint;
        const deterministicVerification = options.deterministicAssertions === undefined
            ? undefined
            : await verifyWorkbookAssertions(result.outputPath, options.deterministicAssertions);
        // Verifier 2.0: deterministic check of the planner's own assertions.
        const assertionCheck = planAssertions.length > 0
            ? await verifyWorkbookAssertions(result.outputPath, planAssertions)
            : undefined;
        let verdict = deterministicVerification === undefined
            ? await options.planner.verify({
                ...planContext,
                path: result.outputPath,
                profileSummary: summarizeProfile(afterProfile),
                validationSummary: t('{count} 个公式异常', { count: afterValidation.anomalies.length }),
                executedPlan: plan,
                executedResult: result,
                cellSnapshot,
            })
            : { achieved: deterministicVerification.achieved, reason: deterministicVerification.reason };
        const anomalyNote = afterValidation.anomalies.length === 0
            ? t('公式无异常')
            : t('仍有 {count} 个公式异常', { count: afterValidation.anomalies.length });
        const deterministicNote = t('{anomalies}；文件{changed}实质变化', {
            anomalies: anomalyNote,
            changed: changed ? t('有') : t('没有'),
        });
        if (!changed || afterValidation.anomalies.length > 0) {
            verdict = { achieved: false, reason: t('{reason}（确定性校验：{note}）', { reason: verdict.reason, note: deterministicNote }) };
        }
        if (assertionCheck && !assertionCheck.achieved) {
            verdict = {
                achieved: false,
                reason: t('{reason}（规划器断言未过 {passed}/{total}：{failures}）', {
                    reason: verdict.reason,
                    passed: assertionCheck.passed,
                    total: assertionCheck.total,
                    failures: listJoin(assertionCheck.failures.slice(0, 3), 'semicolon'),
                }),
            };
        }
        rounds.push({ round, plan, result, verdict, deterministicVerification, planAssertions: assertionCheck });
        currentPath = result.outputPath;
        previousPlan = plan;
        previousResult = result;
        verifierNote = verdict.reason;
        if (verdict.achieved) {
            achieved = true;
            break;
        }
    }
    const finalOutput = options.outPath ?? path.replace(/\.xlsx$/i, '.agent.xlsx');
    await copyFile(currentPath, finalOutput);
    const finalAnomalies = (await validateWorkbookFile(finalOutput)).anomalies.length;
    return { outputPath: finalOutput, rounds, achieved, finalAnomalies };
}
/**
 * Re-throw a round failure with the round's plan attached, keeping the failure
 * kind when there is one. Wrapping a `PlanSchemaError` in a plain `Error` would
 * drop the kind, and the benchmark's taxonomy would then classify a planning
 * failure as an execution failure — the exact confusion the kind exists to stop.
 */
function wrapRoundError(error, message) {
    return error instanceof PlanSchemaError ? new PlanSchemaError(error.kind, message) : new Error(message);
}
function summarizeProfile(profile) {
    return profile.sheets.map((sheet) => {
        const headers = sheet.columns.filter((column) => column.header).map((column) => column.header).slice(0, 8).join(' / ');
        return t('{sheet}：{rows} 行 × {columns} 列', { sheet: sheet.sheet, rows: sheet.dataRows, columns: sheet.columnCount }) + (headers ? t('，表头 {headers}', { headers }) : '');
    }).join(listSeparator('semicolon'));
}
function summarizePlanOps(steps) {
    const parts = [];
    for (const step of steps) {
        for (const operation of step.operations) {
            const record = operation;
            const keys = ['range', 'source', 'target', 'start', 'sheet', 'column', 'groupColumn', 'metrics', 'summaryColumns', 'keys', 'criteria', 'filter', 'cells'];
            const args = keys
                .filter((key) => record[key] !== undefined)
                .map((key) => `${key}=${JSON.stringify(record[key])}`);
            parts.push(`${operation.op}${args.length > 0 ? `(${args.join(',')})` : ''}`);
        }
    }
    return parts.join(' -> ');
}
/**
 * Compact cell snapshot for LLM verification. Per-sheet round-robin sampling:
 * each sheet (especially summary sheets the plan just created) contributes
 * cells, instead of "first 80 cells of sheet 1" which hides the evidence the
 * verifier needs for analysis tasks.
 */
async function cellSnapshotOf(path, limit = 96) {
    const cells = await readWorkbookCells(await readFile(path));
    const bySheet = new Map();
    for (const [id, content] of Object.entries(cells)) {
        const sheet = id.slice(0, Math.max(0, id.lastIndexOf('!')));
        const list = bySheet.get(sheet) ?? [];
        list.push([id, content]);
        bySheet.set(sheet, list);
    }
    const groups = [...bySheet.values()];
    const lines = [];
    let index = 0;
    // Round-robin across sheets so late-created output sheets are visible.
    while (lines.length < limit && groups.some((group) => index < group.length)) {
        for (const group of groups) {
            const entry = group[index];
            if (!entry)
                continue;
            lines.push(`${entry[0]}=${entry[1].slice(0, 60)}`);
            if (lines.length >= limit)
                break;
        }
        index++;
    }
    return lines.join('\n');
}
