# 源码模块

每个文件做什么。这是给改这个项目的人看的清单，不是使用说明——用法见
[README](../README.md)，工具清单也在那里。

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
- `src/operation-schema.ts` — `excel_operate` 的 77 操作严格判别联合 schema，
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
- `src/file-benchmark.ts` + `src/corpus/` — ExcelBench lite：113 个文件级真实
  职场任务（编辑/分析/公式/工作流），运行与指标见 [docs/benchmark.md](docs/benchmark.md)。
- `src/index.ts` — dsh plugin entry exposing 25 tools（理解文件 / 公式体检与修复 /
  操作与编排 / 图表与导出，完整清单见上面的「工具」表）。
- `bundle/` — 可发布 dsh bundle：manifest + cordis.patch.yml + 编译产物。
