/**
 * Local ExcelBench run: builds the realistic corpus, executes every canonical
 * plan, verifies integrity, and prints the aggregate report. Not part of the
 * npm test suite; run directly:
 *   node --test tests/invoke-file-benchmark.ts
 */
import { runFileBenchmark } from '../src/file-benchmark.ts'
import { corpusTasks } from '../src/corpus/index.ts'

const report = await runFileBenchmark(corpusTasks)
console.log(JSON.stringify({
  total: report.total,
  success: report.success,
  successRate: report.successRate,
  meanAccuracy: report.meanAccuracy,
  integrityRate: report.integrityRate,
  categories: report.categories,
  failed: report.tasks.filter((task) => !task.success).map((task) => ({
    id: task.id,
    checks: `${task.checksPassed}/${task.checksTotal}`,
    integrityBefore: task.integrityBefore,
    integrityAfter: task.integrityAfter,
  })),
}, null, 2))
