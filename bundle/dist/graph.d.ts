export interface DependencyEdge {
    from: string;
    to: string;
}
export interface DependencyGraph {
    edges: DependencyEdge[];
    successors: Record<string, string[]>;
    predecessors: Record<string, string[]>;
    cycles: string[][];
}
export interface TraceStep {
    /** Cell id, e.g. "Sheet1!B4". */
    cell: string;
    /** Distance from the origin: 1 = directly connected, 2 = two hops, … */
    depth: number;
}
/**
 * Walk the graph from `cell`, breadth-first, in one direction:
 * `precedents` = the cells it reads, `dependents` = the cells that read it.
 *
 * This is Excel's Trace Precedents / Trace Dependents expressed as *data*. The
 * arrows Excel draws are UI state and never stored in the .xlsx, so reproducing
 * the feature has to mean returning the chain rather than the drawing.
 */
export declare function traceDependencies(graph: DependencyGraph, cell: string, direction: 'precedents' | 'dependents', depth?: number): {
    reached: TraceStep[];
    truncated: boolean;
};
export declare function buildDependencyGraph(formulas: Array<{
    id: string;
    formula: string;
}>): DependencyGraph;
//# sourceMappingURL=graph.d.ts.map