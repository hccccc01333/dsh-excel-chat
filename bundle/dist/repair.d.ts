import { type CellPatch } from './patch.ts';
import { type WorkbookScore } from './score.ts';
import { type ValidationResult } from './validator.ts';
export type RepairAdvisor = (cells: Record<string, string>, result: ValidationResult) => Promise<CellPatch[]>;
export declare function generateRepairs(cells: Record<string, string>, result: ValidationResult): CellPatch[];
export interface RepairResult {
    repairs: CellPatch[];
    llmRepairs: CellPatch[];
    before: ValidationResult;
    after: ValidationResult;
    repairedPath: string;
    oracleScore: WorkbookScore | null;
}
export declare function repairWorkbookFile(path: string, llmAdvisor?: RepairAdvisor, cells?: Record<string, string>, oracleCells?: Record<string, string>): Promise<RepairResult>;
//# sourceMappingURL=repair.d.ts.map