import { type CellPatch } from './patch.ts';
export interface CellDiffEntry {
    id: string;
    kind: 'added' | 'removed' | 'changed';
    oldValue: string | null;
    newValue: string | null;
}
export declare function diffCellMaps(before: Record<string, string>, after: Record<string, string>): CellDiffEntry[];
export declare function diffToPatches(entries: CellDiffEntry[]): CellPatch[];
export declare function diffWorkbookFiles(beforePath: string, afterPath: string): Promise<CellDiffEntry[]>;
export interface PatchLog {
    version: 1;
    createdAt: string;
    sourcePath: string;
    patches: CellPatch[];
}
export declare function writePatchLog(path: string, log: PatchLog): Promise<void>;
export declare function readPatchLog(path: string): Promise<PatchLog>;
export declare function applyPatchLog(inputPath: string, log: PatchLog, outputPath?: string): Promise<void>;
export declare function rollbackPatchLog(path: string, log: PatchLog, outPath?: string): Promise<void>;
//# sourceMappingURL=diff.d.ts.map