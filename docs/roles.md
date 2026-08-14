# 岗位指南：运营 / 产品 / 数分的 Excel 用法

dsh-excel-chat 内置 `preset` 岗位模板，一句话按你的岗位出活。

## 运营（ops）

典型任务：日报/周报、渠道与活动效果、KPI 追踪、环比趋势。

```json
{"op":"preset","role":"ops","source":"订单!A1:F7","groupColumn":"B",
 "metrics":[{"column":"F","function":"sum"}]}
```

产出：排序 + 分类小计 + 动态透视汇总（`订单-运营报表`）+ 自动筛选 + 表头样式 +
冻结首行 + 金额列数据条条件格式 + `#,##0.00` 格式。

对话示例：

> 用 preset 按渠道生成运营报表，金额合计，我看看哪些渠道最好。

## 产品（product）

典型任务：用户分群、留存/漏斗、版本对比、埋点清洗。

```json
{"op":"preset","role":"product","source":"订单!A1:F7","groupColumn":"C",
 "metrics":[{"column":"F","function":"sum"},{"column":"D","function":"count"}]}
```

产出：与运营相同的一键报表（`订单-产品分析`），指标列用色阶条件格式，一眼看出
高/中/低。

对话示例：

> 用 preset 按产品做分析报表，金额合计 + 数量计数。

## 数分（data）

典型任务：多表关联、口径核对、异常筛查、下钻筛选。

```json
{"op":"preset","role":"data","source":"订单!A1:F7","groupColumn":"B",
 "metrics":[{"column":"F","function":"sum"},{"column":"D","function":"count"}],
 "filter":{"column":"B","operator":"eq","value":"华东"}}
```

产出：一键报表（`订单-数据分析`）+ 色阶 + 按条件筛出的明细副本（`订单-筛选`）。

对话示例：

> 用 preset 做数据分析，按区域金额合计 + 数量计数，只看华东的数据。

## 进阶组合

- 跨表补列：先 `set` 写 VLOOKUP/XLOOKUP，再跑 `preset`。
- 同比/环比：在数据表加两列公式（`=B2/B1-1`），再让 `report` 汇总。
- 看趋势：`excel_create_chart` 基于汇总表建折线/柱状图（Windows）。
- 交付安全：`protectSheet` 加密码、`pageSetup` 设打印区域。
