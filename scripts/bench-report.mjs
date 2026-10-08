/**
 * Build the public benchmark table from the raw run artifacts.
 *
 * The point is that anyone can produce a row. A run is a JSON file in
 * `bench-results/` — exactly what `node tests/invoke-llm-benchmark.ts` prints —
 * so adding your own model is dropping a file in and running this script. The
 * table is generated rather than hand-written, which is what stops the numbers in
 * the docs from drifting away from the numbers that were measured.
 *
 *   npm run bench        > bench-results/<date>-<model>.json   # run it
 *   npm run bench:report                                       # rebuild the table
 *   npm run bench:report -- --check                            # verify it is current
 *
 * Two guardrails, both earned from real mistakes:
 *
 * - **A run that never reached the model is not a result.** Without an API key
 *   the runner does not fail fast: it walks all 100 tasks and records every one
 *   as `execution`/0 checks. That looks exactly like a 0% model and would be
 *   published as one, so such a run is rejected here instead.
 * - **A row without provenance is not comparable.** Model, date and replanning
 *   rounds decide what the number means; a missing one is a hard error rather
 *   than a blank cell.
 */
import { readFileSync, writeFileSync, readdirSync, mkdirSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

/**
 * True when this file is the entry point, so importing it (from the test suite)
 * does not run the CLI. Comparing `fileURLToPath(import.meta.url)` against the raw
 * `process.argv[1]` does not work on Windows — one is absolute with backslashes,
 * the other is whatever the shell passed, often relative. That silent mismatch once
 * left a CLI printing nothing while every test still passed, because the tests
 * import the module rather than running it.
 */
function isMain() {
  return process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href
}

const ROOT = fileURLToPath(new URL('..', import.meta.url))
const RESULTS = join(ROOT, 'bench-results')
const DOC = join(ROOT, 'docs', 'benchmark-results.md')
const CSV = join(RESULTS, 'table.csv')

const CATEGORY_LABELS = { editing: '编辑', analysis: '分析', formula: '公式', workflow: '工作流' }
const FAILURE_LABELS = {
  verification: 'Verification', replan: 'Replan', planning: 'Planning',
  argument: 'Argument', execution: 'Execution', intent: 'Intent',
  semantic: 'Semantic', 'tool-selection': 'Tool Selection', other: 'Other',
}

const pct = (value) => `${(value * 100).toFixed(1)}%`.replace(/\.0%$/, '%')

/**
 * Validate one raw run and normalise it for the table. Exported so the test can
 * feed it the shapes that must be rejected without touching the results directory.
 */
export function validateRun(name, raw) {
  const rate = raw.successRate ?? (raw.total ? raw.success / raw.total : undefined)
  const failures = Object.entries(raw.failureBreakdown ?? {}).filter(([, n]) => n > 0)
  const failureTotal = failures.reduce((sum, [, n]) => sum + n, 0)
  const allExecution = failureTotal > 0 && failures.every(([kind]) => kind === 'execution')

  if (rate === undefined) throw new Error(`${name}: no successRate and no success/total`)
  if (!raw.model) throw new Error(`${name}: no model — the row would not be comparable`)
  if (!raw.date) throw new Error(`${name}: no date — the row would not be comparable`)
  // A keyless run walks every task and records execution failures; it must not be
  // published as a model that scored zero.
  if (rate === 0 && allExecution) {
    throw new Error(`${name}: every task failed as "execution" with 0% success — this is what a run with no API key looks like, not a benchmark result`)
  }

  return {
    file: name,
    label: raw.label ?? raw.model,
    model: raw.model,
    provider: raw.provider ?? '',
    date: raw.date,
    rounds: raw.rounds ?? '',
    total: raw.total ?? '',
    success: raw.success ?? '',
    rate,
    meanAccuracy: raw.meanAccuracy ?? null,
    integrityRate: raw.integrityRate ?? null,
    categories: raw.categories ?? {},
    failureBreakdown: raw.failureBreakdown ?? {},
    topFailure: failures.sort((a, b) => b[1] - a[1])[0]?.[0] ?? '',
    provenance: raw.provenance ?? '',
  }
}

function loadRuns() {
  return readdirSync(RESULTS)
    .filter((name) => name.endsWith('.json'))
    .sort()
    .map((name) => validateRun(name, JSON.parse(readFileSync(join(RESULTS, name), 'utf8'))))
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.model.localeCompare(b.model)))
}

function render(runs) {
  const categories = Object.keys(CATEGORY_LABELS).filter((key) => runs.some((run) => run.categories[key]))

  const header = ['模型', '日期', '轮数', '任务成功率', '平均准确率', '完整性率', ...categories.map((k) => CATEGORY_LABELS[k]), '最大失败源']
  const lines = [
    `| ${header.join(' | ')} |`,
    `|${header.map(() => '---').join('|')}|`,
  ]
  for (const run of runs) {
    lines.push(`| ${[
      run.model,
      run.date,
      run.rounds,
      `**${pct(run.rate)}**`,
      run.meanAccuracy === null ? '—' : pct(run.meanAccuracy),
      run.integrityRate === null ? '—' : pct(run.integrityRate),
      ...categories.map((key) => (run.categories[key] ? pct(run.categories[key].successRate) : '—')),
      run.topFailure ? `${FAILURE_LABELS[run.topFailure] ?? run.topFailure} ${run.failureBreakdown[run.topFailure]}` : '—',
    ].join(' | ')} |`)
  }

  const notes = runs
    .filter((run) => run.provenance)
    .map((run) => `- **${run.label}**：${run.provenance}`)

  return `# 评测结果（自动生成）

> 本文件由 \`npm run bench:report\` 从 \`bench-results/*.json\` 生成，**不要手改**。
> 每个 JSON 就是 \`node tests/invoke-llm-benchmark.ts\` 的原始输出（外加 \`label\` / \`date\` /
> \`rounds\` 几个用于比较的字段）。手写的数字会和实测漂移，这张表不会。

## 结果

${lines.join('\n')}

任务成功率 = 全部断言通过且公式异常为 0；平均准确率 = 断言通过比例；
完整性率 = 公式异常为 0 的任务比例。语料共 100 个文件级任务
（编辑 35 / 分析 25 / 公式 22 / 工作流 18），见 [benchmark.md](benchmark.md)。

## 自己跑一行

不需要改代码，只要有任意 OpenAI 兼容端点（**本地 Ollama 也行，零 API 成本、数据不出网**）：

\`\`\`sh
git clone https://github.com/hccccc01333/dsh-excel-chat && cd dsh-excel-chat

# 1. 跑一遍，把原始输出存成一行结果（这就是完整的数据来源）
DEEPSEEK_API_KEY=sk-... LLM_BENCH_ROUNDS=3 \\
  npm run bench > bench-results/$(date +%F)-your-model.json

# 2. 重新生成这张表
npm run bench:report
\`\`\`

本地模型示例（不花钱、不联网）：

\`\`\`sh
DEEPSEEK_API_KEY=ollama DEEPSEEK_BASE_URL=http://localhost:11434/v1 \\
  DEEPSEEK_MODEL=qwen2.5:14b npm run bench > bench-results/$(date +%F)-qwen2.5-14b.json
npm run bench:report
\`\`\`

跑之前先确认语料本身是好的（**不调用任何模型，秒级完成**）：

\`\`\`sh
npm run bench:corpus     # 100/100 语料回归
\`\`\`

**注意**：没配 key 时 runner 不会报错退出，而是把 100 个任务全记成
\`execution\` 失败。所以生成表时会**拒绝**这种「全 execution + 0% 成功率」的结果——
它不是模型得 0 分，是根本没跑到模型。托管端点限流严重时，用
\`LLM_BENCH_OUT\` 开断点续跑逐任务补齐，别让限流污染整轮结果。

## 口径说明

${notes.length ? notes.join('\n') : '- （暂无）'}

**可比性**：换模型、换提示词、换重规划轮数都会改变结果。上表每行都记了模型与轮数；
提示词版本见 [CHANGELOG](../CHANGELOG.md) 与 [benchmark.md](benchmark.md)。
`;
}

function toCsv(runs) {
  const categories = Object.keys(CATEGORY_LABELS).filter((key) => runs.some((run) => run.categories[key]))
  const header = ['file', 'label', 'model', 'provider', 'date', 'rounds', 'total', 'success', 'successRate', 'meanAccuracy', 'integrityRate', ...categories, 'topFailure']
  const rows = runs.map((run) => [
    run.file, run.label, run.model, run.provider, run.date, run.rounds, run.total, run.success,
    run.rate, run.meanAccuracy ?? '', run.integrityRate ?? '',
    ...categories.map((key) => (run.categories[key] ? run.categories[key].successRate : '')),
    run.topFailure,
  ])
  return [header, ...rows].map((row) => row.map((cell) => (typeof cell === 'string' && cell.includes(',') ? `"${cell}"` : cell)).join(',')).join('\n') + '\n'
}

export function build() {
  const runs = loadRuns()
  return { runs, markdown: render(runs), csv: toCsv(runs) }
}

export function main(argv = process.argv.slice(2)) {
  const { runs, markdown, csv } = build()
  if (argv.includes('--check')) {
    const current = (() => { try { return readFileSync(DOC, 'utf8') } catch { return null } })()
    if (current === markdown) {
      console.log(`bench:report --check: the table matches ${runs.length} run artifact(s)`)
      return 0
    }
    console.error('bench:report --check: docs/benchmark-results.md is stale; run `npm run bench:report`')
    return 1
  }
  mkdirSync(RESULTS, { recursive: true })
  writeFileSync(DOC, markdown)
  writeFileSync(CSV, csv)
  console.log(`bench:report: ${runs.length} run(s) -> docs/benchmark-results.md + bench-results/table.csv`)
  for (const run of runs) console.log(`  ${run.date}  ${pct(run.rate).padStart(6)}  ${run.model}`)
  return 0
}

if (isMain()) {
  process.exit(main())
}
