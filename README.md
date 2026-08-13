# VERA — Verified Excel Reasoning Agent

独立项目目录（`D:\vera`），与 deepseek-harness 仓库工作区分离，通过已发布的
`@deepseek-ai/dsh-*` npm 包运行。当前模块：

- P0 Formula Pattern Validator — 确定性静默错误检测，无 LLM。
- Formula IR + Compiler — 语义公式中间表示 → 确定性 Excel 公式。

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
- `src/chart-visual.ts` — Excel COM 图表导出 + 可注入视觉评审（VLM 接口）。
- `src/vision.ts` — `visionTextFromContext`：把 `ctx.attachments` + `ctx.llm` 接成视觉评审。
- `src/deepseek.ts` — DeepSeek chat completions 客户端（读 `DEEPSEEK_API_KEY`），接修复顾问。
- `src/patch.ts` — 最小补丁抽象：apply / revert / 写回 workbook。
- `src/repair.ts` — 从验证结果生成确定性修复，写出 `.repaired.xlsx` 并复验。
- `src/workbook.ts` — ExcelJS-based workbook reader: `.xlsx` → cell-content map, and `validateWorkbookFile(path)`.
- `src/index.ts` — dsh plugin entry exposing five tools（validate / compile / repair / diff / charts）。
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
```

真实模型端到端：

```sh
node tests/invoke-real-llm.ts
```

构建并安装 bundle：

```sh
npm run build:bundle
dsh plugin --profile demo add ./bundle
```

发布 bundle（可选）：

```sh
cd bundle && npm pack
```

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

## Known limitations (P0)

- Formula parsing is a lightweight scanner, not a full grammar: quoted strings are stripped,
  cell-like tokens followed by `(` are treated as function names, and exotic constructs
  (e.g. `1E5` inside an expression before a real cell ref) may still mis-parse.
- Whole-column references (`Sales!$H:$H`) do not produce cell-level dependency edges.
- Range edges are enumerated only up to 10,000 cells; larger ranges contribute start/end edges only.
