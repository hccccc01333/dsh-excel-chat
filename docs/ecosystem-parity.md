# 生态功能复刻对照表

dsh-excel-chat 的定位：把主流 AI + Excel 方案的用户可见功能效果，整合进
DeepSeek Harness 的一个插件里。下表记录每个生态仓库的“核心效果”、我们
复刻到哪一步、以及缺口。

| 生态仓库 | 核心功能效果 | 我们的整合状态 | 缺口 / 下一批 |
|---|---|---|---|
| ChatExcel | 对话改表：写值、公式、筛选、排序、汇总、透视，多轮会话 | `excel_operate` 30+ 操作 + 对话即工具调用，已覆盖；入口由 `excel_menu` 兜底 | 多轮会话间的上下文记忆与自检引导（可增强 advisor 提示） |
| ExcelGenius2 | 多智能体任务分解 + 反思；一键上传→摘要+异常检测 | `excel_autofix` 反思闭环、`excel_insight` 摘要+异常、`excel_task` steps 编排 + goal 模式（LLM 规划→执行→验证→重规划）已有；`report` / `preset` 模板已有 | — |
| SheetMind | 数据清洗、公式、图表、透视、报表、模糊匹配 + 自愈 | 清洗操作（含 `normalizeText` 全角/半角）、`fuzzyMatch`、公式修复、图表、透视已有 | 数字格式统一 |
| Excel-Agent | 多模态表格理解 + 自然语言公式生成 + 整行条件高亮 | `excel_profile` / `excel_insight` / `highlightRows` / `excel_compile_formula` 已有 | 截图/图片理解（接入 VLM），公式自然语言解释 |
| sv-excel-agent | Agent 文件操作 + 图表导出 | 已有（`excel_export_charts` 等，Windows + Excel） | — |
| excel-mcp-server / local-workbook-mcp | 通过 MCP 把 Excel 操作暴露给任意 LLM 客户端 | 未做 | MCP 适配层（若 DeepSeek Harness 支持 MCP 工具，再接入） |
| SpreadsheetBench / Spreadsheet-RL | 真实数据集评测 + RL 训练（outcome reward） | oracle 判分 + 11 个合成 benchmark 已有 | 把公开真实数据集转成 BenchmarkTask，作为 RL reward |
| SheetCopilot | 多步任务规划（从指令到完整操作序列） | `preset`（运营/产品/数分模板）+ `report` 已有 | 更多岗位模板 + 指令级规划缓存 |
| SpreadsheetLLM | 表格结构化编码，省 token | `excel_profile` + `excel_read` `maxRows` 分页已有 | 列级语义标注、按意图只读相关行列 |
| cellm / agent-xlsx / xeli | 单元格内 AI 生成/修复公式 | `excel_compile_formula` + `excel_repair_formulas` + `excel_explain_formula` 已有 | — |
| office-ai-agent | Office 文档联动（Excel/Word/PPT） | `mailMerge`（Excel→Word 式模板）已有 | Word/PPT 生成（超出当前范围，暂缓） |

## 已复刻效果清单（按版本）

- v0.29.0 — SpreadsheetLLM 式大表速览（`excel_profile` + `maxRows` 分页）；
  SheetMind 式一键自愈（`excel_autofix`）。
- v0.30.0 — SheetMind / Excel-Agent 式数据清洗：`dedupeRows`、`fillMissing`、
  `removeEmptyRows`、`removeEmptyColumns`、`trimText`、`changeCase`、
  `splitColumn`。
- v0.31.0 — 交互入口：`excel_menu` 文件感知能力菜单（给文件就出可选方案 +
  示例话术），系统提示注册“业务目标优先、模糊就给选项、先做后改可回滚”
  的交互策略，解决用户不会描述的问题。
- v0.32.0 — 复刻三个生态效果：`excel_insight` 数据摘要+异常检测
  （ExcelGenius2）、`highlightRows` 整行条件高亮（Excel-Agent）、
  `fuzzyMatch` 两表模糊匹配（SheetMind）。
- v0.33.0 — `excel_task` 多步编排（ExcelGenius2 复杂任务一次完成）、
  `excel_explain_formula` 公式白话解释（cellm/xeli）、`normalizeText`
  全角/半角标准化（SheetMind 清洗）。
