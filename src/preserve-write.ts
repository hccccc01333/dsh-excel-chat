/**
 * Write a rewritten workbook without silently dropping pivot tables.
 *
 * ExcelJS cannot round-trip `xl/pivotTables/*` and `xl/pivotCache/*`, so several
 * edit paths strip those parts before loading. The host-side operation path has
 * always put them back; the repair / patch-undo / health-report paths had not,
 * which turned a harmless "write a report into the file" into permanent data
 * loss. One helper keeps the rule in one place.
 */
import { t } from './i18n.ts'
import { validateStructure } from './package-check.ts'
import { applyPreserved, capturePreserved, isEmpty } from './preserve.ts'
import { writeWorkbookSafely, type SafeWriteResult } from './safe-write.ts'

/**
 * @param original  Exact bytes of the file before the rewrite (pivot source).
 * @param rewritten Bytes produced by ExcelJS from the stripped input.
 */
export async function writeWorkbookPreserving(
  original: Uint8Array,
  rewritten: Uint8Array,
  outputPath: string,
): Promise<SafeWriteResult> {
  const preserved = capturePreserved(original)
  const merged = applyPreserved(rewritten, preserved)
  if (!isEmpty(preserved)) {
    // A bad merge produces a package Excel refuses to open — worse than the
    // pivot table it was trying to save. Refuse before touching the target.
    const problems = validateStructure(merged)
    if (problems.length > 0) {
      throw new Error(t('保留原有内容时生成的表格未通过完整性校验，已放弃写入以免损坏文件：{problems}', {
        problems: problems.slice(0, 3).join('; '),
      }))
    }
  }
  return writeWorkbookSafely(outputPath, merged)
}
