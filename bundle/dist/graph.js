import { canonicalCellId, columnToNumber, numberToColumn, parseCellId, parseFormula, } from './formula.js';
const MAX_ENUMERATE_CELLS = 10_000;
export function buildDependencyGraph(formulas) {
    const successors = new Map();
    const edgeKeys = new Set();
    const known = new Set();
    function addEdge(from, to) {
        if (from === to)
            return;
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
