# 评测（ExcelBench lite）

从 v0.34.0 起，评测从“11 个合成公式修复 case”升级为**文件级真实任务语料**：
100 个职场场景（编辑 35 / 分析 25 / 公式 22 / 多步工作流 18），每个任务包含
真实中文数据（订单、价目表、区域汇总、销售/市场双表等）、期望操作序列与
单元格/样式断言。

## 指标

- 任务成功率：全部断言通过且验证/修复后公式异常为 0
- 平均准确率：断言通过比例
- 完整性率：验证/修复后异常为 0 的任务比例
- 修复数：验证器发现异常后被确定性修复器修复的数量

## 运行

```sh
node --test tests/invoke-file-benchmark.ts   # 打印聚合报告
node --test tests/file-benchmark.test.ts     # 语料回归守护（100/100）
node --test tests/invoke-llm-benchmark.ts    # 真实 LLM 规划基准（需 DEEPSEEK_API_KEY）
```

## 真实 LLM 规划基准（基线）

用 goal 模式（DeepSeek `deepseek-chat`，LLM 规划 + LLM 验证，maxRounds=2）跑
100 任务语料，最终以任务断言 + 公式完整性打分：

| 指标 | 数值 |
|---|---|
| 任务成功率 | 48% |
| 平均准确率 | 55.1% |
| 完整性率 | 89% |
| 编辑 | 21/35（60%） |
| 公式 | 13/22（59%） |
| 工作流 | 5/18（28%） |
| 分析 | 9/25（36%） |

失败归因（按频次）：

1. 分析类操作参数结构复杂（metrics/groupColumn/outputSheet），纯文本规划器
   容易写错或漏字段——已加 plan schema 校验/自动修复层，缺失必填字段会明确
   报错并回喂规划器重规划。
2. 验证器误判“已达成”——已并入确定性校验（公式异常数 + 值/样式指纹是否
   实质变化），与 LLM 判断合取。
3. 少量 LLM 输出非法 JSON 或把 fill 当 set 用——引擎层补容错（前缀、别名、
   cells 转字符串、数组包装）。
4. 修复类任务：LLM 倾向重写公式而非依赖确定性 autofix，plan 与目标不一致。

三次迭代数字：47%（基线）→ 49% → 48%（增强后；完整性 79% → 89%，分析类
24% → 36%，引擎异常噪音清零，剩余失败全部是语义/规划质量问题）。

复现：`node --test tests/invoke-llm-benchmark.ts`（每次消耗 DeepSeek API）。

## 失败分类（v0.35 Failure Taxonomy）

从 v0.35 起，每次 LLM 基准除了输出成功率，还会把每个失败任务归入一个
可解释分类（`src/failure-taxonomy.ts`），回答“这 48% 到底输在哪”：

| 分类 | 含义 | 判定依据 |
|---|---|---|
| Intent Error | 理解错用户任务 | 完全没用期望操作，且无通用兜底操作 |
| Semantic Error | 找错列/表/指标 | 期望操作都已执行、参数一致，但断言未过 |
| Planning Error | 步骤缺失/顺序错误 / 计划结构无效 | 空计划、缺关键步骤、planner 报错 |
| Tool Selection Error | 该用专用工具却用通用操作 | 期望 `aggregateReport` 却只用了 `set`/`fill` 等 |
| Argument Error | schema/range/column 参数错误 | 操作名正确但参数与期望不一致 |
| Execution Error | Excel 引擎本身出错 | agent 崩溃且错误非规划类 |
| Verification Error | 没完成却判完成 | 验证器判定达成但最终断言失败 |
| Replan Error | 第一轮失败后没有纠正 | 多轮重规划仍未达成 |

基准输出新增 `failureBreakdown` 分类计数，以及每个失败任务的 `failure`
（分类 + 人类可读细节）。分类是确定性启发式，边界（如 semantic vs
argument）允许人工复核后调整。

### v0.35 首轮实测（分析类 5 任务冒烟，deepseek-chat，maxRounds=3）

启用分类 + 引擎加固后重跑前 5 个分析任务：

| 指标 | 加固前 | 加固后 |
|---|---|---|
| 执行崩溃（Execution） | 3 | 0 |
| Planning | 2 | 1 |
| Verification 误判 | 1 | 3 |
| 完整性率 | 0.4 | 0.8 |

结论：引擎崩溃已清零，剩余分析类失败主要是**验证器把没做完的活判成
完成**（3/5），其次是计划结构不完整（conditionalFormatting 缺 rules）。
这正是后续 Verifier 2.0 / Semantic Layer 的主攻方向。

## 语料结构

`src/corpus/` 下按类别组织，`src/file-benchmark.ts` 负责执行与计分。每个任务
是“构建输入 → 执行期望操作 → 断言输出 + 完整性检查”的可复现场景，新增能力
时把对应场景加进语料，回归由测试守护。

## 路线图

- 100 → 500 真实任务（含图表、透视的 Windows/Excel 用例、更多行业场景）
- LLM planner 基准：已跑通（基线 50%），持续优化提示词/验证器并跟踪数字
- 公开 leaderboard
