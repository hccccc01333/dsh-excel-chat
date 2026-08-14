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
export declare function buildDependencyGraph(formulas: Array<{
    id: string;
    formula: string;
}>): DependencyGraph;
//# sourceMappingURL=graph.d.ts.map