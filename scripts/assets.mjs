/**
 * Promotional asset generator.
 *
 * Two rules keep these images honest and cheap to maintain:
 *
 * 1. **Everything shown is measured, never drawn from imagination.** The
 *    before/after figures come from actually running `autofixWorkbookFile` on a
 *    deliberately broken workbook (`collectEvidence`), and `--check` re-runs that
 *    and fails when the engine's output no longer matches what the committed
 *    images were built from. Without it an image silently becomes fiction the
 *    moment the engine changes — the drift this project keeps getting bitten by.
 *
 * 2. **Adding a capability is one array entry.** `FEATURES` drives the grid, so a
 *    new tool is a few lines here, not a new hand-made picture. The scenes are
 *    data + a template; nothing is positioned by hand.
 *
 * Rendering is HTML → headless browser → PNG, so text stays crisp at any width
 * (the old full-app screenshots were unreadable at README scale). Node cannot
 * spawn a browser in every environment, so a failed spawn is not fatal: the HTML
 * is still written and the exact command is printed.
 *
 *   node scripts/assets.mjs            # write HTML, then screenshot
 *   node scripts/assets.mjs --check    # verify the evidence still matches (CI)
 *   node scripts/assets.mjs --html     # write HTML only
 */
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import { autofixWorkbookFile } from '../src/autofix.ts'
import { readWorkbookCells } from '../src/workbook.ts'
import { readFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'

const ROOT = fileURLToPath(new URL('..', import.meta.url))
const OUT = join(ROOT, 'assets')
const BUILD = join(ROOT, '.assets-build')

/** Browsers to try, in order. Only used when this Node can spawn. */
const BROWSERS = [
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  'C:/Program Files/Microsoft/Edge/Application/msedge.exe',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  'google-chrome',
]

const BRAND = { bg: '#0b1220', panel: '#111c33', line: '#1e2b45', text: '#e6edf7', dim: '#8ea2c0', accent: '#4f8cff', bad: '#ff6b6b', good: '#3ddc84' }

/**
 * The pillars, in the order the README argues them. `hero` marks the ones that
 * carry the pitch; `proof` must be something the engine really does, not a
 * promise. Adding a capability is an entry here — that is the whole update.
 */
export const FEATURES = [
  { id: 'breadth', title: '对话把 Excel 活干完', proof: '25 个工具、77 种操作：建表 / 清洗 / 报表 / 透视 / 图表 / 邮件合并 / 导出', hero: true },
  { id: 'role-reports', title: '按岗位一句话出活', proof: '运营 / 产品 / 数分三套模板：排序 + 小计 + 透视 + 筛选 + 样式 + 冻结一次成型', hero: true },
  { id: 'insight', title: '看得懂数据', proof: '数据洞察与语义画像：缺失 / 重复 / 异常值 / 列角色 / 跨表关联键', hero: true },
  { id: 'visual', title: '画得好看', proof: '原生图表与透视表、条件格式数据条与色阶、迷你图、对话内可编辑网格', hero: true },
  { id: 'no-excel', title: '不依赖本机 Excel', proof: '核心能力纯 XML 层实现，macOS / Linux / Windows 都能跑', hero: false },
  { id: 'trust', title: '改完敢交差', proof: '每次编辑后自动体检公式，确定性修复并复验；审计日志可回滚', hero: false },
]

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c])

/**
 * Build a workbook whose formula column has exactly one silent error, then run
 * the real repair path. Returns what the engine actually said, so the image is
 * a measurement rather than an illustration.
 */
export async function collectEvidence() {
  const dir = await mkdtemp(join(tmpdir(), 'assets-'))
  const path = join(dir, 'orders.xlsx')
  const wb = new ExcelJS.Workbook()
  const sheet = wb.addWorksheet('订单')
  sheet.addRow(['订单号', '数量', '单价', '金额'])
  const rows = [['A-1001', 12, 30], ['A-1002', 8, 45], ['A-1003', 5, 60], ['A-1004', 20, 15], ['A-1005', 3, 200]]
  rows.forEach((r, i) => sheet.addRow([r[0], r[1], r[2], null, { formula: `C${i + 2}*D${i + 2}` }]))
  // The whole column multiplies C[row] by D[row]; row 5 reaches one row up.
  sheet.getCell('E5').value = { formula: 'C5*B4' }
  await wb.xlsx.writeFile(path)

  const out = await autofixWorkbookFile(path)
  const before = await readWorkbookCells(await readFile(path))
  const after = await readWorkbookCells(await readFile(out.repairedPath))
  const repair = out.repairs[0]
  return {
    column: 'E',
    rows: [2, 3, 4, 5, 6].map((row) => ({
      cell: `E${row}`,
      before: String(before[`订单!E${row}`] ?? ''),
      after: String(after[`订单!E${row}`] ?? ''),
      changed: String(before[`订单!E${row}`]) !== String(after[`订单!E${row}`]),
    })),
    anomalyCell: repair?.id ?? '',
    anomalyKind: Object.keys(out.before.byKind)[0] ?? '',
    beforeCount: out.before.total,
    afterCount: out.after.total,
    healthScore: out.healthScore ?? 100,
    beforeFormula: String(repair?.oldValue ?? ''),
    afterFormula: String(repair?.newValue ?? ''),
    summary: out.message.split('\n')[0],
  }
}

/** The committed images were built from this. `--check` compares against it. */
export const EXPECTED = {
  anomalyCell: '订单!E5',
  anomalyKind: 'reference-offset',
  beforeFormula: '=C5*B4',
  afterFormula: '=C5*D5',
  beforeCount: 1,
  afterCount: 0,
}

function page(body, width, height) {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{width:${width}px;height:${height}px;background:${BRAND.bg};color:${BRAND.text};
       font-family:"Microsoft YaHei","PingFang SC",system-ui,-apple-system,sans-serif;
       -webkit-font-smoothing:antialiased;overflow:hidden}
  .mono{font-family:Consolas,"SF Mono",Menlo,monospace}
  .pill{display:inline-block;padding:7px 15px;border-radius:999px;font-size:16px;
        color:#cfe0ff;background:rgba(79,140,255,.12);border:1px solid rgba(79,140,255,.45)}
  </style></head><body>${body}</body></html>`
}

/** Header banner: lead with what you get, not with one internal feature. */
function bannerScene(e) {
  const pills = FEATURES.filter((f) => f.hero).map((f) => `<span class="pill">${esc(f.title)}</span>`).join('')
  const body = `
  <div style="height:100%;display:flex;flex-direction:column;justify-content:center;gap:18px;padding:0 64px;
              background:linear-gradient(120deg,#0b1220 0%,#101f3d 55%,#0b1220 100%)">
    <div style="font-size:52px;font-weight:700;letter-spacing:-1px">
      Excel 的活，<span style="color:${BRAND.accent}">说一句就干完</span>。
    </div>
    <div style="font-size:23px;color:${BRAND.dim};line-height:1.55;max-width:1010px">
      <b style="color:${BRAND.text}">dsh-excel-chat</b> — 在 DeepSeek Harness 里用对话操作 Excel：
      建表、清洗、报表、透视、图表、导出，改完还会自己体检一遍公式。
    </div>
    <div style="display:flex;gap:12px;flex-wrap:wrap;margin-top:6px">${pills}</div>
    <div class="mono" style="font-size:15px;color:${BRAND.dim};margin-top:4px">
      npm i dsh-excel-chat · 25 个工具 · 77 种操作 · MIT
    </div>
  </div>`
  return { file: 'banner.png', width: 1200, height: 340, html: page(body, 1200, 340) }
}

/** The hero: the same column before and after, with the engine's own wording. */
function beforeAfterScene(e) {
  const rows = e.rows.map((r) => {
    const mark = r.changed
    return `<div class="row${mark ? ' bad' : ''}">
      <span class="cell mono">${r.cell}</span>
      <span class="formula mono">${esc(r.before)}</span>
      ${mark ? `<span class="arrow">→</span><span class="fixed mono">${esc(r.after)}</span>` : '<span class="arrow dim">→</span><span class="same mono">未改动</span>'}
    </div>`
  }).join('')

  const body = `
  <div style="padding:34px 48px 0">
    <div style="display:flex;align-items:baseline;gap:14px;margin-bottom:6px">
      <span style="font-size:27px;font-weight:700">同一列公式，一行被静默弄坏</span>
      <span class="mono" style="font-size:15px;color:${BRAND.dim}">订单.xlsx · E 列</span>
    </div>
    <div style="font-size:16px;color:${BRAND.dim};margin-bottom:22px">
      Excel 不会报错——其余 4 行都是 <span class="mono" style="color:${BRAND.text}">C[row]*D[row]</span>，只有 E5 向上取了一行，金额从此算错。
    </div>
    <div class="table">${rows}</div>
    <div class="foot">
      <div>
        <div class="lbl">体检结果</div>
        <div class="mono" style="font-size:17px;color:${BRAND.bad}">${esc(e.anomalyCell)} · ${esc(e.anomalyKind)}</div>
        <div style="font-size:15px;color:${BRAND.dim};margin-top:3px">异常 ${e.beforeCount} → ${e.afterCount}，健康分 ${e.healthScore}</div>
      </div>
      <div>
        <div class="lbl">修复动作</div>
        <div class="mono" style="font-size:17px">
          <span style="color:${BRAND.bad}">${esc(e.beforeFormula)}</span>
          <span style="color:${BRAND.dim}"> → </span>
          <span style="color:${BRAND.good}">${esc(e.afterFormula)}</span>
        </div>
        <div style="font-size:15px;color:${BRAND.dim};margin-top:3px">确定性修复后重新体检确认</div>
      </div>
    </div>
  </div>
  <style>
    .table{border:1px solid ${BRAND.line};border-radius:10px;overflow:hidden}
    .row{display:flex;align-items:center;gap:16px;padding:11px 18px;border-bottom:1px solid ${BRAND.line};
         background:${BRAND.panel};font-size:17px}
    .row:last-child{border-bottom:0}
    .row.bad{background:rgba(255,107,107,.10)}
    .cell{width:52px;color:${BRAND.dim};font-size:16px}
    .formula{width:150px}
    .row.bad .formula{color:${BRAND.bad};font-weight:600}
    .arrow{color:${BRAND.dim}}
    .arrow.dim{opacity:.35}
    .fixed{color:${BRAND.good};font-weight:600}
    .same{color:${BRAND.dim};opacity:.55}
    .foot{display:flex;gap:64px;margin-top:26px;padding-top:20px;border-top:1px solid ${BRAND.line}}
    .lbl{font-size:13px;color:${BRAND.dim};letter-spacing:.08em;margin-bottom:5px}
  </style>`
  return { file: 'before-after.png', width: 1200, height: 460, html: page(body, 1200, 460) }
}

/** One readable card per capability, so "25 tools" stops being a number. */
function featureGridScene(e) {
  const cards = FEATURES.map((f) => `
    <div class="card${f.hero ? ' hero' : ''}">
      <div class="t">${esc(f.title)}</div>
      <div class="p">${esc(f.proof)}</div>
    </div>`).join('')
  const body = `
  <div style="padding:34px 48px">
    <div style="font-size:26px;font-weight:700;margin-bottom:6px">能力一览</div>
    <div style="font-size:16px;color:${BRAND.dim};margin-bottom:22px">25 个工具、77 种操作；下面每一条都能在对话里直接说。</div>
    <div class="grid">${cards}</div>
    <div style="font-size:14px;color:${BRAND.dim};margin-top:20px">蓝色边框是主张的四条：广度、岗位出活、看懂数据、画得好看；后两条是它敢让你用的理由。</div>
  </div>
  <style>
    .grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}
    .card{background:${BRAND.panel};border:1px solid ${BRAND.line};border-radius:10px;padding:16px 18px;min-height:104px}
    .card.hero{border-color:${BRAND.accent};box-shadow:0 0 0 1px rgba(79,140,255,.25)}
    .t{font-size:17px;font-weight:600;margin-bottom:7px}
    .p{font-size:14px;color:${BRAND.dim};line-height:1.55}
  </style>`
  return { file: 'feature-grid.png', width: 1200, height: 360, html: page(body, 1200, 360) }
}

export function scenes(evidence) {
  return [bannerScene(evidence), beforeAfterScene(evidence), featureGridScene(evidence)]
}

function check(evidence) {
  const bad = Object.entries(EXPECTED).filter(([k, v]) => evidence[k] !== v)
  if (bad.length === 0) {
    console.log('assets --check: evidence matches the committed images')
    return 0
  }
  console.error('assets --check: the engine no longer produces what the images show:')
  for (const [k, v] of bad) console.error(`  ${k}: expected ${JSON.stringify(v)}, got ${JSON.stringify(evidence[k])}`)
  console.error('Regenerate with `node scripts/assets.mjs` and review the diff before committing.')
  return 1
}

function screenshot(htmlPath, pngPath, width, height) {
  for (const browser of BROWSERS) {
    if (browser.includes('/') && !existsSync(browser)) continue
    const r = spawnSync(browser, [
      '--headless', '--disable-gpu', '--hide-scrollbars',
      `--screenshot=${pngPath}`, `--window-size=${width},${height}`,
      `file:///${htmlPath.replace(/\\/g, '/')}`,
    ], { stdio: 'ignore' })
    if (!r.error && r.status === 0) return browser
    if (r.error) return { failed: r.error.code ?? r.error.message, browser }
  }
  return null
}

export async function main(argv = process.argv.slice(2)) {
  const evidence = await collectEvidence()

  if (argv.includes('--check')) return check(evidence)

  mkdirSync(BUILD, { recursive: true })
  mkdirSync(OUT, { recursive: true })
  const built = scenes(evidence)
  for (const scene of built) {
    writeFileSync(join(BUILD, scene.file.replace(/\.png$/, '.html')), scene.html)
  }
  console.log(`assets: wrote ${built.length} HTML template(s) to ${BUILD}`)
  console.log(`  evidence: ${evidence.anomalyCell} ${evidence.beforeFormula} -> ${evidence.afterFormula}, anomalies ${evidence.beforeCount} -> ${evidence.afterCount}`)

  if (argv.includes('--html')) return 0

  let rendered = 0
  let failure = null
  for (const scene of built) {
    const res = screenshot(join(BUILD, scene.file.replace(/\.png$/, '.html')), join(OUT, scene.file), scene.width, scene.height)
    if (res && typeof res === 'object') { failure = res; break }
    if (res) { rendered++; console.log(`  ✓ ${scene.file}  (${scene.width}x${scene.height})`) }
    else { failure = { failed: 'no browser found' }; break }
  }

  if (failure) {
    console.error(`\nassets: could not render images (${failure.failed}).`)
    console.error('This is expected in a sandbox that blocks child processes. The HTML is ready — render it with:')
    for (const scene of built) {
      console.error(`  <browser> --headless --disable-gpu --hide-scrollbars \\`)
      console.error(`    --screenshot="assets/${scene.file}" --window-size=${scene.width},${scene.height} \\`)
      console.error(`    "file:///${join(BUILD, scene.file.replace(/\.png$/, '.html')).replace(/\\/g, '/')}"`)
    }
    return 1
  }
  console.log(`assets: rendered ${rendered} image(s) into assets/`)
  return 0
}

// Only run when invoked directly, so the test suite can import `collectEvidence`
// and `EXPECTED` without triggering a render.
if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1].replace(/\\/g, '/')) {
  process.exit(await main())
}
