# Changelog

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
