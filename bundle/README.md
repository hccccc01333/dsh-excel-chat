# dsh-excel-chat

在 DeepSeek Harness 里对话完成 Excel 工作的 dsh 插件。

[![npm version](https://img.shields.io/npm/v/dsh-excel-chat)](https://www.npmjs.com/package/dsh-excel-chat)

## 一分钟上手（使用者）

前提：已安装 DeepSeek Harness（`dsh` CLI 或桌面端）。

```sh
dsh plugin --profile demo add dsh-excel-chat
dsh web --profile demo
```

对话示例：“帮我把 report.xlsx 做成报表：D 列毛利、E 列合计、表头加粗、冻结首行、
加筛选”；“检查 sales.xlsx 的公式哪里错了并修复”；“按区域生成透视表和柱状图”。

平台说明：公式校验/修复、读写、样式、汇总、合并、邮件合并、插入图片跨平台；图表
创建/改参、原生透视表、图表导出、视觉评审、PDF 导出需要 Windows + 本机 Excel。

安装（从 npm 或本地 bundle）：

```sh
dsh plugin --profile demo add dsh-excel-chat
dsh plugin --profile demo add ./bundle
```

## 工具

**理解文件**

- `excel_read` — 精确读取单元格状态（值/公式/类型/格式/合并/数据有效性）
- `excel_profile` — 大表速览：表头、每列类型/缺失/唯一值/区间/样例 + 建议读取范围
- `excel_semantic_profile` — 语义画像：列角色（时间/维度/指标/标识）、粒度、跨表关联键
- `excel_menu` — 不知道怎么描述就给菜单：一句话摘要 + 清洗/报表/透视/图表等可选方案
- `excel_insight` — 数据洞察：摘要 + 缺失/重复/异常值/负值/空格等体检 + 下一步建议
- `excel_preview` — 表格预览：Markdown 表格（对话内展示）+ HTML 预览文件

**公式**

- `excel_validate_formulas` — 公式静默错误检测（列 pattern、hardcode、空行、循环引用、错误值）
- `excel_compile_formula` — Formula IR → 确定性 Excel 公式
- `excel_explain_formula` — 公式白话解释（函数、引用区域、跨表引用）
- `excel_trace` — 公式依赖链路：引用（它读了谁）/ 从属（谁读了它），可指定层数
- `excel_find_errors` — 列出所有错误值单元格，附产生它的公式与按错误码计数
- `excel_repair_formulas` — 确定性修复 + 可选 LLM 修复，输出 `.repaired.xlsx` 并复验
  （可用 `autoTable` 自动识别表头）
- `excel_autofix` — 一键自愈：体检 → 修复 → 复检 → 人话汇报，附带隐藏健康报告表
- `excel_health_report` — 把体检报告写进工作簿本身（隐藏表：健康分 + 异常清单）

**操作与任务**

- `excel_operate` — 职场级 Excel 操作：写值（自动类型识别）、填充/序列、行列增删、
  复制/移动/转置、格式刷、公式转值、排序（多键/按颜色/自定义序列）、`report` 一键
  报表、分类汇总、动态透视、交叉表、两表关联回填、高级筛选、样式、数据有效性、
  条件格式、自动筛选、结构化表格、插入图片、导入/导出 CSV、页面设置、页眉页脚、
  命名区域、冻结窗格、隐藏行列、分组折叠、超链接、批注、迷你图、查找替换、工作表
  保护、邮件合并、数据清洗（去重/补空/分列/大小写/全角半角）、模糊匹配；操作后
  自动复验公式并写审计日志
- `excel_task` — 多步编排（每步自动体检/修复）+ **goal 闭环**（LLM 规划 → 执行 →
  验证 → 未达成自动重规划）
- `excel_undo` — 按审计日志回滚 `excel_operate` 的编辑
- `excel_diff_workbook` — 两个 workbook 的单元格差异

**图表与导出**

- `excel_validate_charts` — 图表结构校验
- `excel_validate_charts_visual` — Excel 导出 PNG + 视觉 LLM 评审（Windows）
- `excel_create_chart` / `excel_modify_chart` — 用本地 Excel 创建与修改图表（类型、
  标题、图例、坐标轴，Windows）
- `excel_create_pivot` — 原生数据透视表（多行/列字段、筛选器 + 值字段，Windows）
- `excel_export_charts` — 用本地 Excel 把图表导出为 PNG（Windows）
- `excel_export_pdf` — 用本地 Excel 把工作簿或单个工作表导出为 PDF（Windows）

构建（源码安装时 `prepare` 自动执行）：

```sh
npm run build
```
