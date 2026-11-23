import { type SafeWriteResult } from './safe-write.ts';
/**
 * @param original  Exact bytes of the file before the rewrite (pivot source).
 * @param rewritten Bytes produced by ExcelJS from the stripped input.
 */
export declare function writeWorkbookPreserving(original: Uint8Array, rewritten: Uint8Array, outputPath: string): Promise<SafeWriteResult>;
//# sourceMappingURL=preserve-write.d.ts.map