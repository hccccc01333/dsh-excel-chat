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
- `src/patch.ts` — 最小补丁抽象：apply / revert / 写回 workbook。
- `src/repair.ts` — 从验证结果生成确定性修复，写出 `.repaired.xlsx` 并复验。
- `src/workbook.ts` — ExcelJS-based workbook reader: `.xlsx` → cell-content map, and `validateWorkbookFile(path)`.
- `src/index.ts` — dsh plugin entry exposing `excel_validate_formulas`, `excel_compile_formula`, and `excel_repair_formulas`.

## Run tests

```sh
node --test tests/formula-validator.test.ts
node --test tests/compiler.test.ts
node --test tests/workbook-reader.test.ts
node --test tests/patch.test.ts
node --test tests/repair.test.ts
node --test tests/ir-schema.test.ts
node --test tests/advisor.test.ts
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
