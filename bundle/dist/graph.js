import { canonicalCellId, columnToNumber, numberToColumn, parseCellId, parseFormula, } from './formula.js';
const MAX_ENUMERATE_CELLS = 10_000;
/**
 * Walk the graph from `cell`, breadth-first, in one direction:
 * `precedents` = the cells it reads, `dependents` = the cells that read it.
 *
 * This is Excel's Trace Precedents / Trace Dependents expressed as *data*. The
 * arrows Excel draws are UI state and never stored in the .xlsx, so reproducing
 * the feature has to mean returning the chain rather than the drawing.
 */
export function traceDependencies(graph, cell, direction, depth = 1) {
    // Edges run from a formula cell to the cells it *reads*, so `successors` are a
    // cell's precedents and `predecessors` are its dependents — the opposite of
    // what the two names suggest at a glance.
    const adjacency = direction === 'precedents' ? graph.successors : graph.predecessors;
    const parsed = parseCellId(cell);
    const origin = canonicalCellId(parsed.sheet, parsed.column, parsed.row);
    const seen = new Set([origin]);
    const reached = [];
    let frontier = [origin];
    let truncated = false;
    for (let level = 1; level <= depth && frontier.length > 0; level++) {
        const next = [];
        for (const node of frontier) {
            for (const neighbour of adjacency[node] ?? []) {
                if (seen.has(neighbour))
                    continue;
                seen.add(neighbour);
                reached.push({ cell: neighbour, depth: level });
                next.push(neighbour);
            }
        }
        // More links exist past the requested depth — flag it so the returned count
        // is never mistaken for the whole chain.
        if (level === depth && next.some((node) => (adjacency[node] ?? []).length > 0))
            truncated = true;
        frontier = next;
    }
    return { reached, truncated };
}
export function buildDependencyGraph(formulas) {
    const successors = new Map();
    const edgeKeys = new Set();
    const known = new Set();
    function addEdge(from, to) {
        // Self-edges are deliberately kept. A formula that reads its own cell —
        // `A1 = A1+1`, or `A1 = SUM(A1:A5)` — is Excel's most common circular
        // reference, and skipping the edge here made it invisible to `findCycles`
        // and therefore to the validator's `circular-reference` anomaly.
        const key = `${from}->${to}`;
        if (edgeKeys.has(key))
            return;
        edgeKeys.add(key);
        if (!successors.has(from))
            successors.set(from, new Set());
        successors.get(from).add(to);
    }
    for (const entry of formulas) {
        let cell;
        try {
            cell = parseCellId(entry.id);
        }
        catch {
            continue;
        }
        const from = canonicalCellId(cell.sheet, cell.column, cell.row);
        known.add(from);
        let parsed;
        try {
            parsed = parseFormula(entry.formula);
        }
        catch {
            continue;
        }
        const baseSheet = cell.sheet;
        for (const ref of parsed.references) {
            for (const point of [ref.start, ref.end]) {
                if (!point || point.row === null)
                    continue;
                addEdge(from, canonicalCellId(point.sheet ?? baseSheet, point.column, point.row));
            }
            if (ref.end && ref.start.row !== null && ref.end.row !== null) {
                const startCol = columnToNumber(ref.start.column);
                const endCol = columnToNumber(ref.end.column);
                const colMin = Math.min(startCol, endCol);
                const colMax = Math.max(startCol, endCol);
                const rowMin = Math.min(ref.start.row, ref.end.row);
                const rowMax = Math.max(ref.start.row, ref.end.row);
                const area = (colMax - colMin + 1) * (rowMax - rowMin + 1);
                if (area <= MAX_ENUMERATE_CELLS) {
                    for (let col = colMin; col <= colMax; col++) {
                        for (let row = rowMin; row <= rowMax; row++) {
                            addEdge(from, canonicalCellId(ref.start.sheet ?? baseSheet, numberToColumn(col), row));
                        }
                    }
                }
            }
        }
    }
    return {
        edges: [...edgeKeys].map((key) => {
            const separator = key.indexOf('->');
            return { from: key.slice(0, separator), to: key.slice(separator + 2) };
        }),
        successors: toRecord(successors),
        predecessors: toRecord(buildPredecessors(successors)),
        cycles: findCycles(known, successors),
    };
}
function buildPredecessors(successors) {
    const predecessors = new Map();
    for (const [from, targets] of successors) {
        for (const target of targets) {
            if (!predecessors.has(target))
                predecessors.set(target, new Set());
            predecessors.get(target).add(from);
        }
    }
    return predecessors;
}
function toRecord(map) {
    const record = {};
    for (const [key, values] of map)
        record[key] = [...values].sort();
    return record;
}
function findCycles(nodes, successors) {
    const color = new Map();
    const stack = [];
    const seen = new Set();
    const cycles = [];
    function visit(node) {
        color.set(node, 1);
        stack.push(node);
        const targets = successors.get(node);
        if (targets) {
            for (const next of targets) {
                if (!nodes.has(next))
                    continue;
                const state = color.get(next) ?? 0;
                if (state === 0) {
                    visit(next);
                }
                else if (state === 1) {
                    const start = stack.indexOf(next);
                    const cycle = [...stack.slice(start), next];
                    const key = [...cycle].sort().join(',');
                    if (!seen.has(key)) {
                        seen.add(key);
                        cycles.push(cycle);
                    }
                }
            }
        }
        stack.pop();
        color.set(node, 2);
    }
    for (const node of nodes) {
        if ((color.get(node) ?? 0) === 0)
            visit(node);
    }
    return cycles;
}
