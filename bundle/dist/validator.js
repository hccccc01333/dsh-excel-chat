import { buildDependencyGraph } from './graph.js';
import { detectEmptyGaps, detectHardcodeBreaks, detectPatternAnomalies, } from './patterns.js';
export function validate(cells) {
    const formulaEntries = [];
    const errors = [];
    for (const [id, content] of Object.entries(cells)) {
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
    const columns = detectPatternAnomalies(cells);
    const anomalies = [
        ...columns.flatMap((column) => column.anomalies),
        ...detectHardcodeBreaks(cells),
        ...detectEmptyGaps(cells),
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
        cellCount: Object.keys(cells).length,
        formulaCount: formulaEntries.length,
        dependencyGraph,
        columns,
        anomalies,
        errors,
    };
}
