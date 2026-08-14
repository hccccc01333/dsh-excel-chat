export interface CellPatch {
    id: string;
    kind: 'formula' | 'value';
    oldValue: string;
    newValue: string;
}
export declare function applyPatches(cells: Record<string, string>, patches: CellPatch[]): Record<string, string>;
export declare function revertPatches(cells: Record<string, string>, patches: CellPatch[]): Record<string, string>;
export declare function applyPatchesToWorkbook(inputPath: string, patches: CellPatch[], outputPath?: string): Promise<void>;
//# sourceMappingURL=patch.d.ts.map