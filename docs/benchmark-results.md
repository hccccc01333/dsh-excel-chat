# 评测结果（自动生成）

> 本文件由 `npm run bench:report` 从 `bench-results/*.json` 生成，**不要手改**。
> 每个 JSON 就是 `node tests/invoke-llm-benchmark.ts` 的原始输出（外加 `label` / `date` /
> `rounds` 几个用于比较的字段）。手写的数字会和实测漂移，这张表不会。

## 结果

| 模型 | 日期 | 轮数 | 任务成功率 | 平均准确率 | 完整性率 | 编辑 | 分析 | 公式 | 工作流 | 最大失败源 |
|---|---|---|---|---|---|---|---|---|---|---|
| deepseek-chat | 2026-10-07 | 3 | **52%** | 61% | 93% | 65.7% | 28% | 68.2% | 38.9% | Verification 33 |
| glm-5.3-flash | 2026-10-07 | 3 | **86%** | 88.5% | 99% | 85.7% | 84% | 95.5% | 77.8% | Argument 9 |

任务成功率 = 全部断言通过且公式异常为 0；平均准确率 = 断言通过比例；
完整性率 = 公式异常为 0 的任务比例。语料共 100 个文件级任务
（编辑 35 / 分析 25 / 公式 22 / 工作流 18），见 [benchmark.md](benchmark.md)。

## 自己跑一行

不需要改代码，只要有任意 OpenAI 兼容端点（**本地 Ollama 也行，零 API 成本、数据不出网**）：

```sh
git clone https://github.com/hccccc01333/dsh-excel-chat && cd dsh-excel-chat

# 1. 跑一遍，把原始输出存成一行结果（这就是完整的数据来源）
DEEPSEEK_API_KEY=sk-... LLM_BENCH_ROUNDS=3 \
  npm run bench > bench-results/$(date +%F)-your-model.json

# 2. 重新生成这张表
npm run bench:report
```

本地模型示例（不花钱、不联网）：

```sh
DEEPSEEK_API_KEY=ollama DEEPSEEK_BASE_URL=http://localhost:11434/v1 \
  DEEPSEEK_MODEL=qwen2.5:14b npm run bench > bench-results/$(date +%F)-qwen2.5-14b.json
npm run bench:report
```

跑之前先确认语料本身是好的（**不调用任何模型，秒级完成**）：

```sh
npm run bench:corpus     # 100/100 语料回归
```

**注意**：没配 key 时 runner 不会报错退出，而是把 100 个任务全记成
`execution` 失败。所以生成表时会**拒绝**这种「全 execution + 0% 成功率」的结果——
它不是模型得 0 分，是根本没跑到模型。托管端点限流严重时，用
`LLM_BENCH_OUT` 开断点续跑逐任务补齐，别让限流污染整轮结果。

## 口径说明

- **DeepSeek 基线（v0.36 前提示词）**：docs/benchmark.md 的 v0.35/v0.37 对比表转录。原始 JSONL 未随仓库提交，所以这一行不能由脚本重新生成。
- **glm-5.3-flash（v0.37 提示词）**：docs/benchmark.md 的 v0.37 全量实测转录。换模型与换提示词同时发生，86% 是两者叠加的结果，不是提示词的单独贡献。托管端点限流严重，该轮依赖 JSONL 断点续跑逐任务补齐。原始 JSONL 未随仓库提交。

**可比性**：换模型、换提示词、换重规划轮数都会改变结果。上表每行都记了模型与轮数；
提示词版本见 [CHANGELOG](../CHANGELOG.md) 与 [benchmark.md](benchmark.md)。
