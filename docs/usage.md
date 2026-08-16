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

装完先自检一次（检查宿主包隔离 + 引擎冒烟，装坏了一跑就知道）：

```sh
dsh-excel-chat-doctor
# 或：~/.dsh/profiles/demo/node_modules/.bin/dsh-excel-chat-doctor
# 对话里也可以直接发 /excel-doctor
```

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
| `excel_semantic_profile` | 语义画像：列角色（时间/维度/指标/标识）、数据粒度、派生指标、跨表关联键；分析类任务先跑它，避免找错列 |
| `excel_menu` | 不会描述也没关系：给文件就出菜单——一句话摘要 + 清洗/报表/透视/图表/体检/通知/岗位模板等可选方案，每个带示例话术，直接选 |
| `excel_insight` | 数据洞察：一句话摘要 + 缺失/重复/异常值/负值/空格等体检 + 下一步建议 |
| `excel_preview` | 表格预览：Markdown 表格（对话内展示）+ HTML 预览文件 |
| `excel_task` | 多步编排（steps 模式，逐步体检/修复）+ Agent 闭环（goal 模式：LLM 规划→执行→验证→重规划） |
| `excel_explain_formula` | 公式白话解释：函数、引用区域、跨表引用、运算 |
| `excel_validate_formulas` | 静默公式错误检测：列 pattern、结构不匹配、硬编码、空行、循环引用、`#REF!` 等错误值 |
| `excel_compile_formula` | 语义 Formula IR（binary/ratio/aggregate/function）→ 确定性公式 |
| `excel_repair_formulas` | 确定性修复 + 可选 LLM 修复，输出修复副本并复验 |
| `excel_autofix` | 一键自愈：体检 → 确定性修复（可选 LLM）→ 复检 → 人话汇报；修复副本自动附带隐藏健康报告表 |
| `excel_health_report` | 把公式体检报告写进工作簿本身（隐藏「_dsh_体检报告」表：健康分 + 异常清单），报告跟着文件走 |
| `excel_operate` | 30+ 种精细化操作：写值、填充/序列、行列增删、复制/移动、排序、分类汇总、动态透视报表、高级筛选、样式（字号/边框）、数据有效性、条件格式（数据条/色阶/图标集）、自动筛选、结构化表格、页面设置、命名区域、冻结窗格、查找替换、工作表保护、邮件合并、工作表管理、合并、数据清洗（去重/填缺失/删空行空列/去空格/大小写/全角半角/分列）、整行条件高亮、两表模糊匹配；自动写审计日志 |
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

### 不知道怎么开口？

> 用户：D:\data\销售台账.xlsx，帮我看看能做啥

agent 会自动调用 `excel_menu`，先给你一句话说清这张表里有什么（行数、列数、
表头、空值、公式），再列出一份菜单：

1. 数据清洗——去重、补空值、去空格、分列
2. 一键经营报表——按区域分组、金额合计
3. 动态透视汇总 / 原生透视表
4. 图表——柱状图 / 折线图 / 饼图
5. 公式体检 + 自愈
6. 批量通知 / 岗位模板（运营 / 产品 / 数分）

每个选项都带一句示例话术，直接回复编号或把示例话术发过来即可；做完不满意
可以用 `excel_undo` 回滚。把“描述需求”变成“选需求”，不需要你会 Excel 术语。

### 数据清洗

> 把 sales.xlsx 按“产品”去重（保留第一条），B 列空值填 0，去掉名称首尾空格，
> 再把 SKU 列按 “-” 分成两列。

一次 `excel_operate` 的 operations 数组完成：`dedupeRows` →
`fillMissing`（mode: value / forward / left）→ `trimText` →
`splitColumn`（自动在右侧插入新列，已有列右移、公式引用联动）。
还有 `removeEmptyRows` / `removeEmptyColumns` / `changeCase`
（upper / lower / proper）。

### 数据洞察、高亮与模糊匹配

> 帮我看看 sales.xlsx 这表有什么问题 / 把订单表和价目表按名称模糊匹配，把编码
> 填回 C 列 / 把“苹果”所在的行全部标黄。

`excel_insight` 一次给出摘要 + 缺失/重复/异常值/负值/空格等体检结论和建议；
`excel_operate` 的 `highlightRows` 按条件整行高亮（默认黄色），`fuzzyMatch`
把两表按键相似度匹配（如对账、名称有错别字）并回填目标列，可同时输出匹配
分数。

### 多步任务与公式解释

> 先清洗 sales.xlsx（去重、补空值），再按区域汇总，一次做完 / 帮我解释一下
> D 列这个公式是什么意思。

`excel_task` 把多个 `excel_operate` 步骤串成一次调用，每步结束自动体检公式、
坏了自动修复，再进入下一步（复杂任务一次完成）；`excel_explain_formula`
把公式翻译成人话：用了哪些函数、引用哪些区域、是否跨表。

### 表格预览

> 看看 D:\demo-sales.xlsx 这张表长什么样。

`excel_preview` 把表渲染成 Markdown 表格直接显示在对话里，同时在工作簿旁生成
`<文件名>.preview.html`，用浏览器打开就能看到完整的表格视图（公式、样式按
单元格内容展示）。

### Goal 模式（Agent 闭环）

> 把这个表做成一张月度经营报表，金额按区域汇总、表头加粗、自动筛选。

不写步骤，直接给目标：`excel_task` 的 `goal` 模式会让 LLM 规划操作序列，
执行后由 LLM 验证器判断目标是否达成；没达成就带着上一轮结果重新规划，
最多 `maxRounds` 轮（默认 2）。每一步仍然经过公式体检 + 确定性修复。

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
