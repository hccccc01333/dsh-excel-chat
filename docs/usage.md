# dsh-excel-chat 使用指南

在 DeepSeek Harness 里用自然语言操作 Excel：读、写、算、画、改，最后自动体检
公式有没有被弄坏。

## 1. 安装

前提：已安装 DeepSeek Harness（`dsh` CLI 或桌面端）。

```sh
# 从 npm 安装（推荐）
dsh plugin --profile demo add dsh-excel-chat

# 或从 GitHub 安装
# dsh plugin --profile demo add github:hccccc01333/dsh-excel-chat

# 打开对话界面
dsh web --profile demo
```

锁定版本：`dsh plugin --profile demo add dsh-excel-chat@0.23.1`（不写版本默认 latest）。

## 2. 一分钟开始

打开对话后，直接说：

- “帮我把 report.xlsx 做成报表：D 列毛利、E 列合计、表头加粗、冻结首行、加筛选”
- “检查 sales.xlsx 的 D 列公式有没有错，不对的修掉”
- “按区域生成透视表，金额合计，再生成柱状图”
- “用模板给每笔订单生成发货通知”

agent 会自动调用工具完成，并把结果文件路径告诉你。

## 3. 工具总览

| 工具 | 作用 |
|---|---|
| `excel_read` | 精确读取：值/公式/类型/数字格式/字体/填充/对齐/合并/数据有效性 |
| `excel_profile` | 大表速览：表头、每列类型/缺失/唯一值/数值区间/高频值/样例 + 建议读取范围；配合 `maxRows` 分页读大表 |
| `excel_validate_formulas` | 静默公式错误检测：列 pattern、结构不匹配、硬编码、空行、循环引用、`#REF!` 等错误值 |
| `excel_compile_formula` | 语义 Formula IR（binary/ratio/aggregate/function）→ 确定性公式 |
| `excel_repair_formulas` | 确定性修复 + 可选 LLM 修复，输出修复副本并复验 |
| `excel_autofix` | 一键自愈：体检 → 确定性修复（可选 LLM）→ 复检 → 人话汇报 |
| `excel_operate` | 30+ 种精细化操作：写值、填充/序列、行列增删、复制/移动、排序、分类汇总、动态透视报表、高级筛选、样式（字号/边框）、数据有效性、条件格式（数据条/色阶/图标集）、自动筛选、结构化表格、页面设置、命名区域、冻结窗格、查找替换、工作表保护、邮件合并、工作表管理、合并；自动写审计日志 |
| `excel_undo` | 按审计日志回滚一次 `excel_operate` 编辑 |
| `excel_diff_workbook` | 两个 workbook 的单元格级差异 |
| `excel_validate_charts` | 图表结构校验 |
| `excel_create_chart` / `excel_modify_chart` | 创建/修改图表（类型/标题/图例/坐标轴），Windows + Excel |
| `excel_create_pivot` | 原生数据透视表（多行/列字段、筛选器 + 值字段），Windows + Excel |
| `excel_export_charts` | 导出图表 PNG，Windows + Excel |

## 4. 典型场景

### 报表搭建

> 帮我把 report.xlsx 做成报表：D 列毛利（收入-成本）、E 列合计、表头加粗填浅灰、
> 冻结第一行、加筛选。

agent 会用一次 `excel_operate` 完成 set + style + freezePanes + autoFilter，并返回
操作后公式复验结果。

### 公式体检与修复

> 检查 sales.xlsx 的 D 列公式是否都是“收入-成本”，不对的修掉，输出到 repaired.xlsx。

流程：`excel_read` 看清现状 → `excel_validate_formulas` 定位异常 →
`excel_repair_formulas` 修复并复验。

### 大表速览与一键自愈

> 先看一下 D:\data\销售台账.xlsx 的结构，再按页读数据。

流程：`excel_profile` 返回紧凑画像（每个表/每列的类型、缺失、唯一值、数值
区间、高频值、样例与建议读取范围）；大表按 `readHint` 起手、用 `excel_read`
的 `maxRows` 分页读取，避免一次灌入整张表。

> 刚才那轮修改后帮我检查并修一下公式。

`excel_autofix` 一条命令完成“体检 → 确定性修复（可选 `useLlm: true` 走 LLM）
→ 复检 → 汇报”，输出修复副本与修复前后的异常数。

### 数据分析（透视 / 汇总 / 高级筛选）

> 按区域生成动态透视报表（金额合计 + 数量计数），再给订单表做分类汇总，F 列加
> 数据条条件格式，最后保护工作表。

`aggregateReport` 生成的是实时 SUMIFS 联动报表；`excel_create_pivot` 生成的是
Excel 原生可交互透视表；`subtotal` 负责分类小计。

想要一条指令出整套经营报表，用 `report` 模板操作（排序 + 分类汇总 + 动态透视
汇总 + 自动筛选 + 表头样式 + 冻结首行 + 数字格式一步完成）：

> 用 report 给订单表按区域生成经营报表：金额合计，输出到“经营报表”，数字格式
> 用 #,##0.00。

### 跨表补列（VLOOKUP）

> 给订单表加一列“产品名称”，用 VLOOKUP 从产品价目表按产品代码查。

### 批量通知（邮件合并）

> 用“通知模板”表（含 {区域}、{数量} 占位符）给每笔订单生成一行通知，放到
> “发货通知”表。

### 图表

> 用 excel_create_chart 给订单表建柱状图：区域为分类、金额为数值，标题“区域金额”。

Windows + Excel 环境可用；图表创建后可继续用 `excel_modify_chart` 改类型/标题/
坐标轴。

## 5. 平台与限制

- 公式校验、修复、读写、样式、汇总、合并、邮件合并等跨平台（macOS / Linux / Windows）。
- 图表创建/改参、原生透视表、图表导出依赖 Windows + 本机安装 Microsoft Excel。
- `excel_undo` 是内容级撤销：单元格值/公式/样式会恢复；插入/删除的行列结构不会
  消失（内容会还原）。如果文件在编辑后被再次修改，回滚会拒绝执行而不是覆盖。
- 删除行列时，引用被删单元格的公式会变成 `#REF!`（与 Excel 行为一致），验证器
  会把它报为错误值。

## 6. 常见问题

**模型找不到我的文件？** 告诉它文件的绝对路径（如 `D:\data\销售台账.xlsx`），或先
让它用 `excel_read` 看一下。

**为什么报告里有异常但没修复？** 有些异常（如结构不匹配）需要 LLM 修复，调用
`excel_repair_formulas` 时加 `useLlm: true`；确定性修复只处理引用偏移和空行填充。

**想一条命令体检 + 修复？** 用 `excel_autofix`：默认走确定性修复；结构不匹配
加 `useLlm: true` 交给 LLM。

**撤销失败说前置条件不匹配？** 说明文件在操作之后又被改过，日志与现状对不上，
这是保护机制，不是 bug。

## 7. 反馈

问题与建议请到 <https://github.com/hccccc01333/dsh-excel-chat/issues> 提交。
