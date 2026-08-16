import { buildDependencyGraph } from './graph.js';
import { detectEmptyGaps, detectErrorValues, detectHardcodeBreaks, detectPatternAnomalies, } from './patterns.js';
/** Skip plugin-owned internal sheets (e.g. `_dsh_体检报告`) so a health
 * report never flags itself or pollutes user-facing validation. */
function isInternalSheetCell(id) {
    const bang = id.lastIndexOf('!');
    if (bang < 0)
        return false;
    return id.slice(0, bang).replace(/^'|'$/g, '').toUpperCase().startsWith('_DSH_');
}
export function validate(cells) {
    const owned = Object.fromEntries(Object.entries(cells).filter(([id]) => !isInternalSheetCell(id)));
    const formulaEntries = [];
    const errors = [];
    for (const [id, content] of Object.entries(owned)) {
        const trimmed = content.trim();
        if (!trimmed.startsWith('='))
            continue;
        formulaEntries.push({ id, formula: trimmed });
    }
    let dependencyGraph;
    try {
        dependencyGraph = buildDependencyGraph(formulaEntries);
    }
    catch (error) {
        errors.push({ id: '*', message: error instanceof Error ? error.message : String(error) });
        dependencyGraph = { edges: [], successors: {}, predecessors: {}, cycles: [] };
    }
    const columns = detectPatternAnomalies(owned);
    const anomalies = [
        ...columns.flatMap((column) => column.anomalies),
        ...detectHardcodeBreaks(owned),
        ...detectEmptyGaps(owned),
        ...detectErrorValues(owned),
    ];
    for (const cycle of dependencyGraph.cycles) {
        anomalies.push({
            kind: 'circular-reference',
            cell: cycle[0] ?? '',
            message: `circular reference: ${cycle.join(' -> ')}`,
            expected: 'acyclic',
            actual: cycle.join(' -> '),
            confidence: 1,
        });
    }
    return {
        cellCount: Object.keys(owned).length,
        formulaCount: formulaEntries.length,
        dependencyGraph,
        columns,
        anomalies,
        errors,
    };
}
