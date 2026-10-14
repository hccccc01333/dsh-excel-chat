/**
 * The five-minute tour: the whole loop, minus the model.
 *
 * "AI 修改表格 → 检测错误 → 修复 → 验证" has two halves. The *conversation* half needs a
 * host and a key. The other half — make the edit, notice the quietly broken formula it
 * turned up, repair it, prove the workbook is clean again — is deterministic and runs
 * right here. Showing that first means a reader can judge the project before installing
 * anything, and it keeps working as the plugin changes, which a screen recording does not.
 *
 * The edit is a real one, through the same dispatcher the tools use: the sentence below is
 * the third thing the README tells a reader to type, and this is what the engine does with
 * it. Nothing here is a mock-up.
 *
 * It runs on a copy in the temp directory, so the file in the repository is never touched,
 * and the copy is removed on the way out. Nothing is left behind.
 *
 *   node scripts/demo.mjs
 */
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { autofixWorkbookFile } from '../src/autofix.ts'
import { operateWorkbookFile } from '../src/operations.ts'
import { readWorkbookCells, validateWorkbookFile } from '../src/workbook.ts'
import { EXAMPLE_PATH, EXAMPLE_SHEET, ORDERS } from './make-example.mjs'

/** The ask, in the words the README suggests. `G` is the derived `金额` column. */
export const DEMO_EDIT = {
  op: 'aggregateReport',
  source: `${EXAMPLE_SHEET}!A1:G${1 + ORDERS.length}`,
  groupColumn: 'C',
  metrics: [{ column: 'G', function: 'sum' }],
}

export const SUMMARY_SHEET = `${EXAMPLE_SHEET}-汇总`

/**
 * Run the loop on files inside `workdir`, and return what each step produced.
 *
 * Split out from the rendering so the test can assert on the same steps the reader sees,
 * rather than on a re-implementation of them that can drift from what is printed.
 *
 * @param {string} workdir
 * @returns {Promise<{
 *   edit: import('../src/operations.ts').OperateResult,
 *   groups: string[],
 *   repairs: Array<{ id: string, oldValue: string, newValue: string }>,
 *   repairedPath: string,
 *   final: import('../src/workbook.ts').ValidationResult,
 * }>}
 */
export async function runTour(workdir) {
  const working = join(workdir, 'working.xlsx')
  await copyFile(EXAMPLE_PATH, working)

  // ① Edit. The same call `excel_operate` makes; it writes an audit log beside the output.
  const edit = await operateWorkbookFile(working, [DEMO_EDIT], join(workdir, 'edited.xlsx'))

  const cells = await readWorkbookCells(await readFile(edit.outputPath))
  const groups = Object.keys(cells)
    .filter((id) => id.startsWith(`${SUMMARY_SHEET}!A`) && id !== `${SUMMARY_SHEET}!A1`)
    .sort((a, b) => Number(a.replace(/\D/g, '')) - Number(b.replace(/\D/g, '')))
    .map((id) => cells[id])

  // ② Detect is `edit.validation` — the check the engine runs after every edit on its own.
  // ③ Repair.
  const fix = await autofixWorkbookFile(edit.outputPath, {
    outPath: join(workdir, 'repaired.xlsx'),
    healthReport: false,
  })

  // ④ Verify by reading the repaired file back.
  const final = await validateWorkbookFile(fix.repairedPath)

  return {
    edit,
    groups,
    repairs: fix.repairs.map((patch) => ({ id: patch.id, oldValue: patch.oldValue, newValue: patch.newValue })),
    repairedPath: fix.repairedPath,
    final,
  }
}

const baseName = (path) => path.split(/[\\/]/).pop()
const rule = (title) => `\n${title}`

/**
 * The text the reader sees, as a pure function of the run.
 *
 * Kept separate from `main` so the test asserts on the rendered output — a change that
 * quietly drops a stage, or makes the walkthrough stop demonstrating anything, shows up as
 * a failing assertion rather than as a silent regression in a file nobody re-reads.
 *
 * @param {Awaited<ReturnType<typeof runTour>>} tour
 * @returns {string}
 */
export function renderTour(tour) {
  const lines = []
  lines.push('dsh-excel-chat —— 一句话改表，然后自己检查有没有改坏')
  lines.push('（不需要模型，也不需要先装 dsh）')
  lines.push(`\n示例文件：examples/sales-ledger.xlsx —— ${ORDERS.length} 条编造的订单，没有真实业务数据`)
  lines.push('演示跑在临时副本上，仓库里的文件不会被改动。')

  // ① Edit.
  lines.push(rule('① 改表   「按区域汇总金额，输出到新表」'))
  lines.push(`   ${SUMMARY_SHEET}：A1「区域」 B1「金额 合计」（表头加粗）`)
  lines.push(`   ${tour.groups.join(' / ')}，B 列是 SUMIFS 实时公式，打开时由 Excel 计算`)
  lines.push(`   审计日志：${baseName(tour.edit.patchLog)}（可回滚）`)

  // ② Detect — the check the engine ran on its own after the edit above.
  const check = tour.edit.validation
  lines.push(rule('② 体检   编辑后自动体检'))
  lines.push(`   ${check.cellCount} 个单元格，${check.formulaCount} 条公式 —— 检出 ${check.anomalies.length} 处异常`)
  for (const anomaly of check.anomalies) {
    lines.push(`   · ${anomaly.cell}  [${anomaly.kind}]`)
    lines.push(`       ${anomaly.message}`)
  }
  lines.push('   （这处错误表里本来就有 —— 你只让它汇总，它顺手查了出来）')

  // ③ Repair.
  lines.push(rule('③ 修复'))
  if (tour.repairs.length === 0) {
    lines.push('   没有可确定性修复的异常。')
  }
  for (const patch of tour.repairs) {
    lines.push(`   ${patch.id}   ${patch.oldValue}  →  ${patch.newValue}`)
  }

  // ④ Verify.
  lines.push(rule('④ 复验'))
  lines.push(`   异常 ${tour.final.anomalies.length} —— 公式列重新一致`)

  // ⑤ Output.
  lines.push(rule('⑤ 输出'))
  lines.push(`   ${baseName(tour.repairedPath)} —— 改前改后都在审计日志里，可回滚`)

  lines.push('\n' + '─'.repeat(64))
  lines.push('这五步都不需要模型：能确定的事就地修掉；不能确定的（比如有人手打一个数字')
  lines.push('覆盖了公式）会明说「未自动修复」，而不是猜一个值填进去。')
  lines.push('\n接上 dsh 和模型之后，同一条闭环从「你说一句话」开始 ——')
  lines.push('你说「按区域汇总金额」，它调上面那个操作，然后自己体检、修复、复验。')
  lines.push('步骤见 README「5 分钟上手」。')
  return lines.join('\n')
}

async function main() {
  const dir = await mkdtemp(join(tmpdir(), 'dsh-excel-chat-demo-'))
  try {
    console.log(renderTour(await runTour(dir)))
  } finally {
    await rm(dir, { recursive: true, force: true }).catch(() => {})
  }
}

// Run only when invoked directly; the test imports the functions and must not print.
if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  await main()
}
