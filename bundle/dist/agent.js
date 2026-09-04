import { copyFile, mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { sanitizePlan } from './plan-schema.js';
import { profileWorkbook } from './profile.js';
import { buildWorkbookSemanticProfile } from './semantic.js';
import { runExcelTask } from './task.js';
import { verifyWorkbookAssertions } from './verifier.js';
import { readWorkbookCells, stripPivotTableParts, validateWorkbookFile } from './workbook.js';
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
            validationSummary: `${beforeValidation.anomalies.length} 个公式异常`,
            previousPlan,
            previousResult,
            verifierNote,
        };
        let plan;
        try {
            plan = await options.planner.plan(planContext);
            if (!plan || plan.length === 0)
                throw new Error('planner returned an empty plan');
            plan = sanitizePlan(plan, planContext.sheetNames).steps;
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            if (round < maxRounds) {
                verifierNote = `计划无效：${message}。请修正后重新规划。`;
                previousPlan = undefined;
                previousResult = undefined;
                continue;
            }
            throw new Error(`${message}（第 ${round} 轮计划：${summarizePlanOps(plan ?? [])}）`);
        }
        const roundOut = join(dir, `round-${round}.xlsx`);
        let result;
        try {
            result = await runExcelTask(currentPath, plan, roundOut);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            if (round < maxRounds) {
                verifierNote = `执行出错：${message}。请修正计划后重新规划。`;
                previousPlan = plan;
                previousResult = undefined;
                continue;
            }
            throw new Error(`${message}（第 ${round} 轮计划：${summarizePlanOps(plan)}）`);
        }
        const afterProfile = await profileWorkbook(result.outputPath);
        const afterValidation = await validateWorkbookFile(result.outputPath);
        const cellSnapshot = await cellSnapshotOf(result.outputPath);
        const changed = (await workbookFingerprint(result.outputPath)) !== beforeFingerprint;
        const deterministicVerification = options.deterministicAssertions === undefined
            ? undefined
            : await verifyWorkbookAssertions(result.outputPath, options.deterministicAssertions);
        let verdict = deterministicVerification === undefined
            ? await options.planner.verify({
                ...planContext,
                path: result.outputPath,
                profileSummary: summarizeProfile(afterProfile),
                validationSummary: `${afterValidation.anomalies.length} 个公式异常`,
                executedPlan: plan,
                executedResult: result,
                cellSnapshot,
            })
            : { achieved: deterministicVerification.achieved, reason: deterministicVerification.reason };
        const deterministicNote = `${afterValidation.anomalies.length === 0 ? '公式无异常' : `仍有 ${afterValidation.anomalies.length} 个公式异常`}；文件${changed ? '有' : '没有'}实质变化`;
        if (!changed || afterValidation.anomalies.length > 0) {
            verdict = { achieved: false, reason: `${verdict.reason}（确定性校验：${deterministicNote}）` };
        }
        rounds.push({ round, plan, result, verdict, deterministicVerification });
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
function summarizeProfile(profile) {
    return profile.sheets.map((sheet) => {
        const headers = sheet.columns.filter((column) => column.header).map((column) => column.header).slice(0, 8).join(' / ');
        return `${sheet.sheet}：${sheet.dataRows} 行 × ${sheet.columnCount} 列${headers ? `，表头 ${headers}` : ''}`;
    }).join('；');
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
async function workbookFingerprint(path) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(stripPivotTableParts(await readFile(path)));
    const parts = [];
    workbook.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                const fill = cell.fill?.type === 'pattern' ? `|fill=${String(cell.fill.fgColor?.argb ?? '')}` : '';
                const bold = cell.font?.bold ? '|bold' : '';
                const numFmt = cell.numFmt && cell.numFmt !== 'General' ? `|fmt=${cell.numFmt}` : '';
                const value = cell.formula ? `=${cell.formula}` : cell.value instanceof Date ? cell.value.toISOString() : String(cell.value ?? '');
                parts.push(`${sheet.name}!${cell.address}=${value}${bold}${numFmt}${fill}`);
            });
        });
    });
    return parts.sort().join('|');
}
