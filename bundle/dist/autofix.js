import { repairWorkbookFile } from './repair.js';
/** Collapse a validation result into counts the model can reason about. */
export function summarizeValidation(result) {
    const byKind = {};
    for (const anomaly of result.anomalies) {
        byKind[anomaly.kind] = (byKind[anomaly.kind] ?? 0) + 1;
    }
    return {
        total: result.anomalies.length,
        byKind,
        formulaCount: result.formulaCount,
        cellCount: result.cellCount,
    };
}
/**
 * One-call self-healing loop: validate, apply deterministic repairs (plus an
 * optional LLM advisor), re-validate the repaired copy, and report a compact
 * before/after summary in plain language.
 */
export async function autofixWorkbookFile(path, options = {}) {
    const result = await repairWorkbookFile(path, options.advisor, undefined, undefined, options.outPath);
    const before = summarizeValidation(result.before);
    const after = summarizeValidation(result.after);
    const repairedPath = result.repairedPath;
    const fixed = result.repairs.length + result.llmRepairs.length;
    const lines = [
        `体检：修复前 ${before.total} 个异常，修复后 ${after.total} 个。`,
    ];
    if (fixed > 0) {
        const ids = [...result.repairs, ...result.llmRepairs].map((patch) => patch.id).join(', ');
        lines.push(`已修复 ${fixed} 处：${ids}`);
    }
    else if (before.total > 0) {
        lines.push(`未自动修复 ${before.total} 处；可提供表格结构后启用 LLM 修复，或人工核对。`);
    }
    else {
        lines.push('未发现公式异常，无需修复。');
    }
    lines.push(`输出文件：${repairedPath}`);
    return { repairedPath, repairs: result.repairs, before, after, message: lines.join('\n') };
}
