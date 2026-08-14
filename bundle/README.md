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

平台说明：公式校验/修复、读写、样式、汇总、合并、邮件合并跨平台；图表创建/改参、
原生透视表、图表导出需要 Windows + 本机 Excel。

安装（从 npm 或本地 bundle）：

```sh
dsh plugin --profile demo add dsh-excel-chat
dsh plugin --profile demo add ./bundle
```

工具：

- `excel_validate_formulas` — 公式静默错误检测（列 pattern、hardcode、空行、循环引用、错误值）
- `excel_compile_formula` — Formula IR → 确定性 Excel 公式
- `excel_read` — 精确读取单元格状态（值/公式/类型/格式/合并/数据有效性）
- `excel_undo` — 按审计日志回滚 `excel_operate` 的编辑
- `excel_repair_formulas` — 确定性修复 + 可选 LLM 修复，输出 `.repaired.xlsx` 并复验
  （可用 `autoTable` 自动识别表头）
- `excel_diff_workbook` — 两个 workbook 的单元格差异
- `excel_operate` — 职场级 Excel 操作（写值自动类型识别、填充/序列、插入/删除
  行列、复制/移动、排序、样式、数据有效性、条件格式、自动筛选、结构化表格、
  冻结、查找替换、工作表管理、合并），操作后自动复验公式
- `excel_validate_charts` — 图表结构校验
- `excel_create_chart` / `excel_modify_chart` — 用本地 Excel 创建与修改图表（类型、
  标题、图例、坐标轴，Windows）
- `excel_create_pivot` — 原生数据透视表（多行/列字段、筛选器 + 值字段，Windows）
- `excel_export_charts` — 用本地 Excel 把图表导出为 PNG（Windows）

构建（源码安装时 `prepare` 自动执行）：

```sh
npm run build
```
