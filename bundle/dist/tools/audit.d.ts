/**
 * Checking a workbook: chart validation, diffing, formula repair, health report.
 *
 * These are the tools that find problems rather than fix them silently, which is
 * why they are separate from the operations. Split out of `index.ts` unchanged.
 */
import type { Context } from '@deepseek-ai/cordis';
export declare function registerAuditTools(ctx: Context): void;
//# sourceMappingURL=audit.d.ts.map