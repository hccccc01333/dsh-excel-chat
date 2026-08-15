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
| 任务成功率 | 50% |
| 平均准确率 | 55.5% |
| 完整性率 | 85% |
| 编辑 | 23/35（66%） |
| 公式 | 14/22（64%） |
| 工作流 | 7/18（39%） |
| 分析 | 6/25（24%） |

失败归因（按频次）：

1. 分析类操作参数结构复杂（metrics/groupColumn/outputSheet），纯文本规划器
   容易写错或漏字段——下一步做 plan schema 校验/自动修复层。
2. 验证器仍有约 1/3 的失败任务误判“已达成”——把确定性校验（公式异常数 +
   是否有实质变化）并入验证。
3. 少量 LLM 输出非法 JSON 或把 fill 当 set 用——引擎层再补容错。
4. 修复类任务：LLM 倾向重写公式而非依赖确定性 autofix，plan 与目标不一致。

复现：`node --test tests/invoke-llm-benchmark.ts`（每次消耗 DeepSeek API）。

## 语料结构

`src/corpus/` 下按类别组织，`src/file-benchmark.ts` 负责执行与计分。每个任务
是“构建输入 → 执行期望操作 → 断言输出 + 完整性检查”的可复现场景，新增能力
时把对应场景加进语料，回归由测试守护。

## 路线图

- 100 → 500 真实任务（含图表、透视的 Windows/Excel 用例、更多行业场景）
- LLM planner 基准：已跑通（基线 50%），持续优化提示词/验证器并跟踪数字
- 公开 leaderboard
