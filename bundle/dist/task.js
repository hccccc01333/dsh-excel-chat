import { copyFile, mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { autofixWorkbookFile } from './autofix.js';
import { applyOperationsToWorkbook } from './operations.js';
import { validateWorkbookFile } from './workbook.js';
/**
 * Multi-step Excel workflow in one call (ExcelGenius2/SheetCopilot-style task
 * orchestration): each step applies an operations array, validates formulas,
 * auto-repairs anomalies with the deterministic fixer, and feeds the verified
 * result into the next step. Intermediate files stay in a temp dir; only the
 * final output is copied to outPath.
 */
export async function runExcelTask(path, steps, outPath) {
    if (steps.length === 0)
        throw new Error('task requires at least one step');
    const dir = await mkdtemp(join(tmpdir(), 'vera-task-'));
    let current = path;
    const results = [];
    for (const [index, step] of steps.entries()) {
        const stepOut = index === steps.length - 1 ? join(dir, 'final.xlsx') : join(dir, `step-${index + 1}.xlsx`);
        const applied = await applyOperationsToWorkbook(current, step.operations, stepOut);
        const result = { index, name: step.name ?? `step ${index + 1}`, warnings: applied.warnings };
        if (step.verify !== false) {
            const validation = await validateWorkbookFile(stepOut);
            if (validation.anomalies.length > 0) {
                const fix = await autofixWorkbookFile(stepOut);
                result.validation = { before: validation.anomalies.length, after: fix.after.total, fixed: fix.repairs.length };
                current = fix.repairs.length > 0 ? fix.repairedPath : stepOut;
            }
            else {
                result.validation = { before: 0, after: 0, fixed: 0 };
                current = stepOut;
            }
        }
        else {
            current = stepOut;
        }
        results.push(result);
    }
    const finalOutput = outPath ?? path.replace(/\.xlsx$/i, '.task.xlsx');
    await copyFile(current, finalOutput);
    const finalAnomalies = (await validateWorkbookFile(finalOutput)).anomalies.length;
    return { outputPath: finalOutput, steps: results, finalAnomalies };
}
