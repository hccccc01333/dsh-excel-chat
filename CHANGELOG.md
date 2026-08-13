# Changelog

## v0.10.0 — 2026-08-14

- 范围尾引用修复：确定性 repair 现在覆盖 `=SUM(B4:C3)` 这类 range.end 偏移异常；
  两个端点同时偏移时一次重建整段范围，未偏移端点保留原文本（含 `$` 绝对修饰）。
- 自动表头检测：`detectTableFromCells` 从单元格内容推断 `{ sheet, columns }`；
  `excel_repair_formulas` 新增 `autoTable` 参数，`useLlm` 时无需手写 table schema。
- Oracle 判分：`scoreWorkbookAgainstOracle` 按单元格对比候选与标准 workbook，
  容忍公式大小写/空白与数字格式差异，输出准确率与 mismatch 明细。
- Pass@1 Benchmark：`runBenchmark` 按“确定性修复 → LLM 修复”真实流程执行任务，
  与 oracle 对比给出 Pass@1 与平均准确率；`tests/invoke-benchmark.ts` 可接真实 DeepSeek。
- 测试规模 56+ 增至 75+，覆盖自动表头、范围修复、判分与 benchmark。

## v0.9.0 — 2026-08-14

- 包名确定为 `dsh-excel-vera-plugin`（dsh 生态前缀 + Excel 品牌），
  `cordis.patch.yml` 与 bundle README 同步更新，发布前改名。

## v0.8.0 — 2026-08-14

- 真实 DeepSeek 模型端到端：`deepseekChatCompletion` / `deepseekLlmTextFromEnv` 直接调用
  chat completions（读 `DEEPSEEK_API_KEY`），接到修复顾问后完成 读取 → 验证 → LLM 生成 IR →
  编译 → Patch → 复验 全流程；`invoke-real-llm` 已用真实 API 跑通
  （`Sheet1!D3: =SUM(B3:C3) → =B3-C3`，复验异常归零）。
- 结构异常检测改进：只标记缺少“多数派槽位”的单元格，D2/D4 正常、D3 异常时不再误报多数派。
- LLM 输出容错：operand 裸字符串自动归一化为 cell/column；裸单元格 id 自动匹配
  sheet 限定键（`D3` → `Sheet1!D3`）；prompt 内置 IR 示例和“只修真正偏离的单元格”约束。

## v0.7.0 — 2026-08-14

- 视觉评审接入真实模型链路：`visionTextFromContext` 用 `ctx.attachments` 上传图表 PNG、
  构造带图片块的用户消息、走 `ctx.llm.stream`；新工具 `excel_validate_charts_visual`。
- 测试用假 attachment store + 假 adapter 验证完整链路（图片块确实随消息发出）。

## v0.6.0 — 2026-08-14

- Chart Visual Validator 骨架：`exportChartsWithExcel` 用本地 Excel COM 把图表导出为 PNG；
  `createVisionCritic` 把视觉 checklist（标题截断 / 图例遮挡 / 标签重叠 / 坐标轴 / 拥挤 / 趋势可读性）
  交给可注入的视觉函数，假实现可完整测试。新工具 `excel_export_charts`。
- bundle 发布准备：MIT LICENSE、bundle README、移除 private；`npm pack --dry-run` 验证
  tarball 包含 `dist/`、`cordis.patch.yml`、README、LICENSE。

## v0.5.0 — 2026-08-14

- Chart Semantic Validator：直接解析 xlsx 内的 chart XML（fflate），结构校验包括
  图表类型、系列引用存在性、缺失单元格、二维范围、日期未排序；新工具 `excel_validate_charts`。
- 可发布 bundle：`bundle/` 目录包含 `dsh.bundle` 清单、`cordis.patch.yml` 和 tsc 编译产物，
  `dsh plugin --profile <name> add ./bundle` 可直接安装；git 安装时 `prepare` 自动构建。
- 全部工具汇总：`excel_validate_formulas` / `excel_compile_formula` / `excel_repair_formulas` /
  `excel_diff_workbook` / `excel_validate_charts`。

## v0.4.0 — 2026-08-14

- `ctx.llm` 接线：`llmTextFromContext` 把 dsh 流式 LLM 服务包装成修复顾问的 `LlmText`，
  `excel_repair_formulas` 新增 `useLlm` / `provider` / `model` / `table` 参数，
  可选注入（`ctx.get('llm')`），无 LLM 服务时确定性修复照常工作。
- Workbook Diff：`diffCellMaps` / `diffWorkbookFiles` 输出 added / removed / changed 单元格，
  新工具 `excel_diff_workbook`（给 Excel 的 Git Diff）。
- Patch Log：`writePatchLog` / `readPatchLog` / `applyPatchLog` / `rollbackPatchLog`，
  修复可审计、可回滚。
- LLM 接线测试使用假 adapter 走完整 `ctx.llm.stream` 管线，无需 API key。

## v0.3.0 — 2026-08-14

- Formula IR JSON Schema：`excel_compile_formula` 的 `ir` 参数改用严格 `oneOf` schema，
  模型输出非法 IR 时在参数校验阶段即被拒绝。
- LLM 修复顾问：`createLlmRepairAdvisor` 把 workbook 摘录 + 异常列表 + 表结构组装成 prompt，
  让 LLM 返回 IR 修复，再走 Compiler → Patch → 复验；LLM 以函数注入，无 API key 也可完整测试。

## v0.2.0 — 2026-08-14

- Formula Patch 抽象：apply / revert / 写回 workbook，带前置条件校验。
- 确定性 Repair：reference-offset 异常 → 生成最小补丁（如 `Sales!D4: =B4-C3 → =B4-C4`）。
- 新工具 `excel_repair_formulas`：读取文件 → 验证 → 修复 → 写出 `.repaired.xlsx` → 复验。
- 公式解析器记录引用文本范围（`range`），支持安全替换。

## v0.1.0 — 2026-08-14

- P0 Formula Pattern Validator：A1 引用解析、依赖图、列 pattern 异常检测、hardcode / 空行 / 循环引用检测。
- Formula IR + Compiler：binary / ratio / aggregate 语义 IR → 确定性 Excel 公式。
- ExcelJS workbook 读取，真实 .xlsx 文件端到端验证。
- dsh 插件入口：`excel_validate_formulas`、`excel_compile_formula`。
