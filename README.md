# dsh-excel-chat — 和 Excel 对话，把活干完

[![npm version](https://img.shields.io/npm/v/dsh-excel-chat)](https://www.npmjs.com/package/dsh-excel-chat)
[![GitHub release](https://img.shields.io/github/v/release/hccccc01333/dsh-excel-chat)](https://github.com/hccccc01333/dsh-excel-chat/releases)
[![license](https://img.shields.io/github/license/hccccc01333/dsh-excel-chat)](LICENSE)

![dsh-excel-chat banner](assets/banner.png)

在 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) 里用自然语言
操作 Excel：说一句“给 D 列加毛利公式、表头加粗、冻结首行、加筛选”，agent 会自动
调用 `excel_operate` 完成；每次编辑后自动体检公式有没有被弄坏，也可以让它
“检查这个表哪里算错了”并自动修复。所有工作都在对话里完成，不需要记 Excel 操作。

![dsh-excel-chat 真实演示（DeepSeek Harness Web + 真实模型录制）](assets/demo.gif)

## 功能实录（对话内真实截图）

| 能力菜单：给文件就给你选项 | 数据洞察：自动发现数据问题 |
| --- | --- |
| ![excel_menu 能力菜单](assets/feature-menu.png) | ![excel_insight 数据洞察](assets/feature-insight.png) |

| 表格预览：对话内真网格 | 公式体检 + 自动修复：差异一目了然 |
| --- | --- |
| ![excel_preview 表格预览](assets/feature-preview.png) | ![excel_autofix 自动修复](assets/feature-autofix.png) |

四张截图都是 DeepSeek Harness Web 里真实模型调用工具后，在消息流工具行内
渲染出的结果：能力菜单、问题清单、可编辑表格、修复前后差异。

## 架构

```mermaid
flowchart LR
  U[用户自然语言] --> H[DeepSeek Harness]
  H --> P["excel_profile / excel_read · 结构速览 / 分页读取"]
  P --> M["excel_menu / excel_insight · 能力菜单 / 数据洞察"]
  M --> O["excel_operate / excel_task · 操作 DSL / 多步编排 / Goal 闭环"]
  O --> V["excel_validate_formulas · 公式体检"]
  V -->|异常| R["excel_autofix / excel_repair_formulas · 确定性修复 + LLM 修复"]
  V -->|干净| OUT[输出 workbook]
  R --> V2[复验]
  V2 --> OUT
  OUT --> X["excel_explain_formula / excel_diff_workbook / excel_undo · 解释 / 对比 / 可回滚"]
```

核心闭环：**理解 → 操作 → 验证 → 修复 → 复验 → 输出**；`excel_task` 的 goal 模式
把这条闭环升级为 **Plan → Act → Observe → Verify → Replan** 的 Agent 循环。

## 安装

```sh
dsh plugin --profile demo add dsh-excel-chat          # 从 npm 安装
dsh plugin --profile demo add ./bundle                # 或本地 bundle 目录
```

装完先自检一次，确认宿主包隔离和引擎都正常：

```sh
dsh-excel-chat-doctor                                  # npm 全局/npx 可用时
# 或 profile 内直接跑：
# ~/.dsh/profiles/demo/node_modules/.bin/dsh-excel-chat-doctor
```

装完直接聊，例如：

> 帮我把 report.xlsx 做成报表：D 列是毛利（收入减成本），E 列加合计，
> 表头加粗填浅灰，冻结第一行，加筛选。

> 检查 sales.xlsx 里 D 列公式是不是每行都是“收入-成本”，不对的帮我修掉。

完整使用指南见 [docs/usage.md](docs/usage.md)，岗位用法（运营/产品/数分）见
[docs/roles.md](docs/roles.md)。

## 给使用者：一分钟上手

前提：已安装 DeepSeek Harness（`dsh` CLI 或桌面端）。

```sh
dsh plugin --profile demo add dsh-excel-chat      # 从 npm 安装
# 或从 GitHub 安装：
# dsh plugin --profile demo add github:hccccc01333/dsh-excel-chat
dsh web --profile demo                             # 打开对话界面
```

然后在对话里直接说：

- “帮我把 report.xlsx 做成报表：D 列毛利、E 列合计、表头加粗、冻结首行、加筛选”
- “检查 sales.xlsx 的 D 列公式有没有错，不对的修掉”
- “按区域生成透视表，金额合计，再生成柱状图”

平台说明：公式校验/修复、读写单元格、样式、汇总、合并、邮件合并等功能跨平台；
图表创建/改参、原生透视表、图表 PNG 导出、PDF 导出需要 Windows + 本机安装 Excel。

锁定版本：`dsh plugin --profile demo add dsh-excel-chat@0.23.0`（不写版本默认 latest）。

## 工具

| 工具 | 作用 |
|---|---|
| `excel_validate_formulas` | 静默公式错误检测：列 pattern 偏移、结构不匹配、hardcode、空行、循环引用、`#REF!`/`#DIV/0!` 等错误值 |
| `excel_compile_formula` | Formula IR（binary / ratio / aggregate / function：VLOOKUP、IF、XLOOKUP、统计、日期等）→ 确定性 Excel 公式 |
| `excel_read` | 精确读取：值/公式/类型/数字格式/字体/填充/对齐/合并/数据有效性，编辑前看清单元格状态 |
| `excel_profile` | 大表速览：识别表头、每列类型/缺失/唯一值/数值区间/高频值/样例，给出建议读取范围；配合 `excel_read` 的 `maxRows` 分页，避免整表灌入对话爆 token |
| `excel_semantic_profile` | 语义画像：把每列分类为 时间/维度/指标/标识，识别数据粒度、派生指标（公式）和跨表关联键；分析类任务先跑它，agent 不再猜“地区是不是 B 列” |
| `excel_menu` | 不会描述也没关系：给文件就能拿到菜单——一句话总结表里有什么，再列出清洗/补空值/报表/透视/图表/体检/通知/岗位模板等可选方案，每个带示例话术，直接选就行 |
| `excel_insight` | 数据洞察：一句话摘要 + 缺失/重复/异常值/负值/空格/公式等启发式体检 + 下一步建议，回答“这表有什么问题”“帮我总结一下” |
| `excel_preview` | 表格预览：把指定表/区域渲染成 Markdown 表格（对话内直接看到）+ HTML 预览文件，回答“看看这个表长什么样” |
| `excel_task` | 两种模式：`steps` 多步编排（每步自动体检公式、坏了自动修）；`goal` Agent 闭环（LLM 规划步骤 → 执行 → 验证 → 未达成自动重规划，最多 maxRounds 轮） |
| `excel_explain_formula` | 公式白话解释：解析函数（SUMIFS/VLOOKUP/IF/日期/文本/统计）、引用区域、跨表引用，回答“这个公式是什么意思” |
| `excel_undo` | 按 `excel_operate` 自动生成的 `.patch.json` 审计日志回滚编辑 |
| `excel_repair_formulas` | 确定性修复 + 可选 LLM 修复（`useLlm` / `autoTable` / `oraclePath` / `outPath`），输出修复副本并复验 |
| `excel_autofix` | 一键自愈闭环：体检 → 确定性修复（可选 LLM）→ 复检 → 人话汇报，输出修复副本（自动附带隐藏健康报告表，可 `healthReport:false` 关闭） |
| `excel_health_report` | 把公式体检报告写进工作簿本身：隐藏「_dsh_体检报告」表，含健康分、异常清单、生成时间，报告跟着文件走 |
| `excel_diff_workbook` | 两个 workbook 的单元格级 diff |
| `excel_operate` | 精细化 Excel 操作：写值、填充/序列、行列增删、复制/移动/转置/仅粘贴值、格式刷（copyStyle）、公式转值（freezeFormulas）、唯一值提取（uniqueValues）、排名列（rankColumn）、排序、`report` 一键报表模板（排序+汇总+动态透视+筛选+样式+冻结+格式）、分类汇总、动态透视报表、二维交叉透视表（crosstab）、两表精确关联回填（joinSheets，无公式 VLOOKUP）、高级筛选、样式（字号/字体/边框/删除线/旋转/缩进）、数据有效性、条件格式（数据条/色阶/图标集）、自动筛选、结构化表格、页面设置、页眉页脚（headerFooter）、打印分页符（rowPageBreaks）、打印标题行/列（printTitles）、命名区域、冻结/取消冻结窗格、缩放（setZoom）、网格线开关（showGridLines）、隐藏行列（hideRows/hideColumns）、行列分组折叠（groupRows/groupColumns）、自适应列宽（autoFitColumnWidths）、超链接（站内跳转与外部 URL）、单元格批注（addComment）、每行趋势迷你图（addSparklines）、查找替换、工作表保护（细化权限）、邮件合并、工作表管理（增删改名复制隐藏标签色/重排 moveSheet）、文档属性与打开时重算（setWorkbookProperties）、合并、取消全部合并（unmergeAll）、数据清洗（去重/填充缺失/删空行空列/去空格/大小写转换/全角半角标准化/分列/区域清除 clearRange）、整行条件高亮（highlightRows）、两表模糊匹配（fuzzyMatch）；操作后自动复验公式并写审计日志 |
| `excel_validate_charts` | 图表结构校验：类型、系列、缺失单元格、二维范围、日期排序 |
| `excel_validate_charts_visual` | Excel 导出 PNG + 视觉 LLM 评审 |
| `excel_export_charts` | 用本地 Excel 把图表导出为 PNG（Windows） |
| `excel_create_chart` | 用本地 Excel 创建图表：数据范围、类型、标题（Windows） |
| `excel_modify_chart` | 修改图表参数：类型、标题、图例、坐标轴（Windows） |
| `excel_create_pivot` | 原生数据透视表（pivotCache + pivotTable）：多行字段、列字段、报表筛选器 + 值字段（求和/计数/平均/最大/最小），Excel 生成、可刷新（Windows） |
| `excel_export_pdf` | 用本机 Excel COM 把工作簿或单个工作表导出为 PDF（Windows，只读打开不动源文件） |

能力深度与可靠性进展：100 个职场任务的自建评测语料（ExcelBench lite）与
真实 LLM 基线见 [docs/benchmark.md](docs/benchmark.md)；右侧可编辑 Excel
面板的设计与实测见 [docs/web-panel.md](docs/web-panel.md)。

## Modules

- `src/formula.ts` — A1 reference parser (cell, range, cross-sheet, whole-column), canonical cell ids, column helpers.
- `src/graph.ts` — dependency graph with bounded range expansion and cycle detection.
- `src/patterns.ts` — per-column reference-pattern analysis: offset anomalies, structure mismatches, hardcode breaks, empty gaps.
- `src/validator.ts` — `validate(cells)` entry point returning graph + column reports + anomalies.
- `src/ir.ts` — Formula IR 类型（binary / ratio / aggregate）。
- `src/ir-schema.ts` — Formula IR 的 dsh 工具 DSL schema（严格 oneOf 校验）。
- `src/compiler.ts` — `compileFormula(ir, { baseCell, table })` 编译为 Excel 公式。
- `src/advisor.ts` — LLM 修复顾问：异常 + 表结构 → prompt → IR 修复 → Patch。
- `src/llm.ts` — `llmTextFromContext`：把 `ctx.llm` 流式服务接入修复顾问（可选注入）。
- `src/diff.ts` — Workbook Diff 与 Patch Log：diff / apply / rollback。
- `src/charts.ts` / `src/chart-validator.ts` — xlsx 图表 XML 解析与结构校验。
- `src/chart-visual.ts` — Excel COM 图表创建/参数修改/导出 + 可注入视觉评审（VLM 接口）。
- `src/vision.ts` — `visionTextFromContext`：把 `ctx.attachments` + `ctx.llm` 接成视觉评审。
- `src/deepseek.ts` — DeepSeek chat completions 客户端（读 `DEEPSEEK_API_KEY`），接修复顾问。
- `src/patch.ts` — 最小补丁抽象：apply / revert / 写回 workbook。
- `src/repair.ts` — 从验证结果生成确定性修复（引用偏移 + 空行填充），写出
  `.repaired.xlsx` 并复验；可选传入 oracle cells 返回 `oracleScore`。
- `src/workbook.ts` — ExcelJS-based workbook reader: `.xlsx` → cell-content map, and `validateWorkbookFile(path)`.
- `src/tables.ts` — `detectTableFromCells`：从单元格内容推断 `{ sheet, columns }`，
  供 `excel_repair_formulas` 的 `autoTable` 自动识别表头。
- `src/score.ts` — `scoreWorkbookAgainstOracle`：oracle 单元格级判分，容忍公式
  大小写/空白与数字格式差异，输出准确率与 mismatch 明细。
- `src/read.ts` — `readWorkbookDetail`：精确读取单元格（值/公式/类型/格式/合并/
  数据有效性），供 `excel_read` 工具使用。
- `src/profile.ts` — `profileWorkbook`：结构化表格编码，输出每表/每列的
  紧凑画像与建议读取范围，供 `excel_profile` 工具使用。
- `src/autofix.ts` — `autofixWorkbookFile`：体检 → 修复 → 复检 → 人话总结的
  一键自愈闭环，供 `excel_autofix` 工具使用。
- `src/pivot.ts` — `createPivotTable`：驱动 Excel COM 生成原生数据透视表
  （pivotCache + pivotTable），保证文件始终合法可打开。
- `src/operation-schema.ts` — `excel_operate` 的 27 操作严格判别联合 schema，
  让模型按 `op` 字段直接生成正确结构。
- `src/operations.ts` — Excel 操作 DSL：set（自动类型识别）/ fill / fillSeries /
  insertRows / deleteRows / insertColumns / deleteColumns（公式引用联动，含跨表，
  被删单元格引用转 `#REF!`）/ sortRange（多键排序）/ copyRange / moveRange /
  style / dataValidation（下拉与数值校验）/ conditionalFormatting / setColumnWidth /
  autoFilter / addTable（结构化表格）/ setRowHeight / freezePanes / findReplace /
  addSheet / renameSheet / deleteSheet / duplicateSheet / hideSheet / setTabColor /
  clear / merge / unmerge。
- `src/benchmark.ts` — Pass@1 benchmark：确定性修复 → LLM 修复，与 oracle 对比判分。
- `src/benchmark-cases.ts` — 11 个 benchmark 任务：范围端点、绝对引用、空行、
  跨表、多表、聚合结构、hardcode 等场景。
- `src/file-benchmark.ts` + `src/corpus/` — ExcelBench lite：100 个文件级真实
  职场任务（编辑/分析/公式/工作流），运行与指标见 [docs/benchmark.md](docs/benchmark.md)。
- `src/index.ts` — dsh plugin entry exposing eight tools（validate / compile / repair / diff / operate / chart structure / chart export / chart visual）。
- `bundle/` — 可发布 dsh bundle：manifest + cordis.patch.yml + 编译产物。

## Run tests

```sh
node --test tests/formula-validator.test.ts
node --test tests/compiler.test.ts
node --test tests/workbook-reader.test.ts
node --test tests/patch.test.ts
node --test tests/repair.test.ts
node --test tests/ir-schema.test.ts
node --test tests/advisor.test.ts
node --test tests/llm-wiring.test.ts
node --test tests/diff.test.ts
node --test tests/chart-validator.test.ts
node --test tests/load-bundle.test.ts
node --test tests/chart-visual.test.ts
node --test tests/pack-bundle.test.ts
node --test tests/vision-wiring.test.ts
node --test tests/deepseek.test.ts
node --test tests/tables.test.ts
node --test tests/score.test.ts
node --test tests/benchmark.test.ts
```

真实模型端到端：

```sh
node tests/invoke-real-llm.ts
node tests/invoke-conversation.ts   # 对话直用：自然语言 -> 工具调用 -> 执行 -> 复验
```

Pass@1 benchmark（确定性修复 + 可选 LLM）：

```sh
node tests/invoke-benchmark.ts                 # 仅确定性路线
VERA_BENCH_LLM=1 node tests/invoke-benchmark.ts # 接真实 DeepSeek
```

构建并安装 bundle：

```sh
npm run build:bundle
dsh plugin --profile demo add ./bundle
```

发布 bundle（可选，本地打包）：

```sh
cd bundle && npm pack
```

自动发布（GitHub Actions，需在仓库配置 `NPM_TOKEN` 自动化 token）：

```sh
git tag v0.18.0 && git push origin v0.18.0
```

CI 会执行测试、构建、`npm pack` 校验 tag 与版本一致、发布 npm，并在 GitHub
Release 上附带 tarball。

通过真实执行管线调用工具：

```sh
node --import tsx tests/invoke-plugin.ts
node --import tsx tests/invoke-compiler.ts
node --import tsx tests/invoke-workbook.ts
node --import tsx tests/invoke-repair.ts
```

挂进 Web UI（可选，两种方式）：

```sh
# 方式一：从仓库目录启动，patch 指向本目录
pnpm dsh web --patch D:/vera/cordis.yml

# 方式二：安装官方 CLI 后从本目录启动（依赖树较大，机器空闲时再装）
npm install --save-dev @deepseek-ai/dsh@0.1.0-rc.6
npx dsh web --patch D:/vera/cordis.yml
```

Windows 上 `cordis.yml` 的入口路径必须是 `file:///D:/vera/src/index.ts` 形式的 URL。

## Example

`D4 = B4-C3` inside a column where every other row is `=B[row]-C[row]` is reported as a
`reference-offset` anomaly with confidence = majority support fraction (e.g. 3/4 = 0.75).

The tool accepts either `cells` (a map) or `path` (an absolute `.xlsx` path) — exactly one.

## 相关链接

- npm：<https://www.npmjs.com/package/dsh-excel-chat>
- GitHub：<https://github.com/hccccc01333/dsh-excel-chat>
- 社区收录：[awesome-dsh-plugin](https://github.com/awesome-dsh-plugin/awesome-dsh-plugin)

## Known limitations (P0)

- Formula parsing is a lightweight scanner, not a full grammar: quoted strings are stripped,
  cell-like tokens followed by `(` are treated as function names, and exotic constructs
  (e.g. `1E5` inside an expression before a real cell ref) may still mis-parse.
- Whole-column references (`Sales!$H:$H`) do not produce cell-level dependency edges.
- Range edges are enumerated only up to 10,000 cells; larger ranges contribute start/end edges only.
