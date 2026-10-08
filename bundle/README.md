# dsh-excel-chat — 和 Excel 对话，把活干完

[![npm version](https://img.shields.io/npm/v/dsh-excel-chat)](https://www.npmjs.com/package/dsh-excel-chat)
[![GitHub release](https://img.shields.io/github/v/release/hccccc01333/dsh-excel-chat)](https://github.com/hccccc01333/dsh-excel-chat/releases)
[![license](https://img.shields.io/github/license/hccccc01333/dsh-excel-chat)](https://github.com/hccccc01333/dsh-excel-chat/blob/master/LICENSE)

![dsh-excel-chat](https://raw.githubusercontent.com/hccccc01333/dsh-excel-chat/master/assets/banner.png)

在 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) 里用自然语言
操作 Excel：说一句「给 D 列加毛利公式、表头加粗、冻结首行、加筛选」，agent 会自动
调用 `excel_operate` 完成；**每次编辑后自动体检公式有没有被弄坏**，也可以让它
「检查这个表哪里算错了」并确定性修复。全程在对话里完成，不用记 Excel 的操作路径。

## 两个不一样的地方

**一、编辑后自动体检，而不是改完就交差。**
Excel 最烦的不是不会写公式，而是公式被**静默弄坏**——拖拽填充把范围截短、插入行后
引用错位、查找失败返回 `#N/A`，Excel 不报错，只是数字开始不对。本插件在每次
`excel_operate` / `excel_task` 之后自动跑一遍公式体检（列内 pattern 一致性、硬编码、
空行断链、循环引用、错误值），发现问题可以用确定性修复器改回来，并把修复前后差异
摆给你看。`excel_undo` 还能按审计日志整体回滚。

**二、核心能力不依赖本机装 Excel。**
公式校验/修复、读写、样式、汇总、合并、邮件合并、插入图片都是**纯 XML 层实现**，
在 macOS / Linux / Windows 上都能跑。只有图表创建改参、原生数据透视表、图表导出
PNG、视觉评审、PDF 导出需要 Windows + 本机 Excel（走 COM）。

## 安装

前提：已安装 DeepSeek Harness（`dsh` CLI 或桌面端）。

```sh
dsh plugin --profile demo add dsh-excel-chat
dsh web --profile demo
```

从源码或本地目录安装也可以（`prepare` 会自动构建）：

```sh
dsh plugin --profile demo add ./bundle
```

### 对话里的快捷命令

| 命令 | 作用 |
| --- | --- |
| `/excel-set` | 直接设置某个单元格的值（绕过规划器，最快路径） |
| `/excel-undo` | 按审计日志回滚上一次 `excel_operate` 编辑 |
| `/excel-doctor` | 自检：插件是否加载、工具注册数、bundle 与当前 dsh 是否兼容 |

工具没出现时先跑自检——它会把「插件没被加载」和「加载了但注册失败」区分开：

```sh
dsh-excel-chat-doctor                                    # npm 全局 / npx 可用时
~/.dsh/profiles/demo/node_modules/.bin/dsh-excel-chat-doctor
```

## 可以这么说

- 「帮我把 report.xlsx 做成报表：D 列毛利、E 列合计、表头加粗、冻结首行、加筛选」
- 「检查 sales.xlsx 的公式哪里错了并修复」
- 「按区域生成透视表和柱状图」
- 「这张表有什么问题？」——会先跑数据洞察 + 公式体检，再给下一步建议
- 「先给我看看这个文件」——输出列类型/缺失/唯一值/样例，以及建议读取范围

## 可靠性

100 个职场任务的[自建评测语料](https://github.com/hccccc01333/dsh-excel-chat/blob/master/docs/benchmark.md)
（ExcelBench lite：编辑 / 分析 / 公式 / 多步工作流），goal 模式 + glm-5.3-flash
全量实测任务成功率 **86%**（DeepSeek 基线 52%）；失败归因、复测方式与已知局限
都写在评测文档里。单元测试 423 项全绿。

## 工具（25 个）

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
  条件格式、自动筛选、结构化表格、插入图片、导入/导出 CSV（导出默认加公式注入
  防护，导入会还原）、页面设置、页眉页脚、命名区域、冻结窗格、隐藏行列、分组折叠、
  超链接、批注、迷你图、查找替换、工作表保护、邮件合并、数据清洗（去重/补空/分列/
  大小写/全角半角）、模糊匹配；操作后自动复验公式并写审计日志
- `excel_task` — 多步编排（每步自动体检/修复）+ **goal 闭环**（LLM 规划 → 执行 →
  验证 → 未达成自动重规划）
- `excel_undo` — 按审计日志回滚 `excel_operate` 的编辑（单元格类型一并还原）
- `excel_diff_workbook` — 两个 workbook 的单元格差异

**图表与导出**

- `excel_validate_charts` — 图表结构校验
- `excel_validate_charts_visual` — Excel 导出 PNG + 视觉 LLM 评审（Windows）
- `excel_create_chart` / `excel_modify_chart` — 用本地 Excel 创建与修改图表（类型、
  标题、图例、坐标轴，Windows）
- `excel_create_pivot` — 原生数据透视表（多行/列字段、筛选器 + 值字段，Windows）
- `excel_export_charts` — 用本地 Excel 把图表导出为 PNG（Windows）
- `excel_export_pdf` — 用本地 Excel 把工作簿或单个工作表导出为 PDF（Windows）

## 实际效果

| 公式体检 + 自动修复：差异一目了然 | 数据洞察：自动发现数据问题 |
| --- | --- |
| ![excel_autofix](https://raw.githubusercontent.com/hccccc01333/dsh-excel-chat/master/assets/feature-autofix.png) | ![excel_insight](https://raw.githubusercontent.com/hccccc01333/dsh-excel-chat/master/assets/feature-insight.png) |

截图都是 DeepSeek Harness Web 里真实模型调用工具后、在消息流工具行内渲染出的结果。
更多截图、演示动图与架构说明见
[GitHub 仓库首页](https://github.com/hccccc01333/dsh-excel-chat)。

## 文档

- [完整使用指南](https://github.com/hccccc01333/dsh-excel-chat/blob/master/docs/usage.md)（场景、参数、平台限制、常见问题）
- [评测方法与结果](https://github.com/hccccc01333/dsh-excel-chat/blob/master/docs/benchmark.md)
- [更新日志](https://github.com/hccccc01333/dsh-excel-chat/blob/master/CHANGELOG.md)

## 开发

```sh
npm run build      # 源码安装时 prepare 会自动执行
node --test tests/*.test.ts
```

License: MIT
