# 评测结果（自动生成）

> 本文件由 `npm run bench:report` 从 `bench-results/*.json` 生成，**不要手改**。
> 每个 JSON 就是 `node tests/invoke-llm-benchmark.ts` 的原始输出（外加 `label` / `date` /
> `rounds` 几个用于比较的字段）。手写的数字会和实测漂移，这张表不会。

## 结果

> ⚠️ **这两行是 100 任务语料时代的数据，不能和现在的语料直接比较。**
> 语料已经从 100 涨到 **144 个任务**（编辑 77 / 分析 27 / 公式 22 / 工作流 18），
> 新加的任务大多来自此前**完全没有覆盖**的操作族 —— 也就是说，这些百分比是在
> **更容易的题面**上得到的。当前语料的实测结果**待重跑**，跑完会替换这一节，
> 并附上原始逐任务结果。

| 模型 | 日期 | 轮数 | 任务成功率 | 平均准确率 | 完整性率 | 编辑 | 分析 | 公式 | 工作流 | 最大失败源 |
|---|---|---|---|---|---|---|---|---|---|---|
| deepseek-chat | 2026-10-07 | 3 | **52%** | 61% | 93% | 65.7% | 28% | 68.2% | 38.9% | Verification 33 |
| glm-5.3-flash | 2026-10-07 | 3 | **86%** | 88.5% | 99% | 85.7% | 84% | 95.5% | 77.8% | Argument 9 |

**上面两行的语料规模是 100 个文件级任务**（编辑 35 / 分析 25 / 公式 22 / 工作流 18）。
任务成功率 = 全部断言通过且公式异常为 0；平均准确率 = 断言通过比例；
完整性率 = 公式异常为 0 的任务比例。当前语料见 [benchmark.md](benchmark.md)。

## 复现

重跑当前语料（144 个任务）。**先跑不花钱的语料自检**，确认题目本身是好的：

```sh
npm run bench:corpus     # 不调用任何模型，秒级完成
```

然后接任意 OpenAI 兼容端点。例如 OpenCode Go 上免费的 Step 5 Preview：

```sh
LLM_PROVIDER=deepseek \
DEEPSEEK_BASE_URL=https://opencode.ai/zen/go/v1 \
DEEPSEEK_MODEL=step-5-preview-free \
DEEPSEEK_API_KEY=<你的 OpenCode key> \
LLM_BENCH_OUT=bench-results/step-5-preview-free.jsonl \
node --test tests/invoke-llm-benchmark.ts
```

`LLM_BENCH_OUT` 是**逐任务断点文件** ✓ —— 每完成一个任务就追加一行 ✓，重跑会跳过已完成的 ✓。
长跑中途断了不用从头来 ✓，而且它**就是**原始逐任务结果 ✓。

**先小样验证再接长跑** ✓：

```sh
LLM_BENCH_SAMPLE=3 LLM_BENCH_OUT=/tmp/smoke.jsonl \
LLM_PROVIDER=deepseek DEEPSEEK_BASE_URL=https://opencode.ai/zen/go/v1 \
DEEPSEEK_MODEL=step-5-preview-free DEEPSEEK_API_KEY=<key> \
node --test tests/invoke-llm-benchmark.ts
```

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
