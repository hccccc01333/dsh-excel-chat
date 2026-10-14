# 贡献指南

谢谢愿意花时间。这个项目对「改得对」的兴趣大于「加得多」——如果一条改动让工具变多、
但让结果变得不可信，它多半会被要求重做。下面是你需要知道的全部。

## 快速开始

```sh
git clone https://github.com/hccccc01333/dsh-excel-chat && cd dsh-excel-chat
npm install

npm test        # 全量测试（直接跑 .ts，不需要先构建）
npm run demo    # 在示例文件上走一遍完整闭环，不需要模型、不需要联网
```

`npm run demo` 是最好的起点：它在 `examples/sales-ledger.xlsx` 的临时副本上跑
**改表 → 体检 → 修复 → 复验 → 输出**，打印引擎的真实输出。你对「这个项目做什么」的
第一个印象，应该来自它，而不是来自文档。

**Node 版本**：测试直接执行 `.ts` 文件（靠 Node 内置的类型擦除，不是编译），所以需要
**Node ≥ 22.18**。CI 跑 Node 24。

注意类型擦除是 **strip-only** 的：`enum`、`namespace`、参数属性、装饰器都会让整个测试
文件加载失败。`tsc` 仍然是绿的，所以症状是「测试总数突然变少」，不是报错。

## 项目结构

| 目录 | 是什么 |
| --- | --- |
| `src/op/` | 13 个操作模块（core / structure / sorting / formatting / filters / reports / editing / layout / csv / images / joins / crosstab / sheets），依赖严格单向、无环 |
| `src/tools/` | 7 个模块，各导出 `registerXxxTools(ctx)`；`src/index.ts` 只做装配 |
| `src/operation-schema.ts` | **`excel_operate` 操作清单的唯一权威** |
| `src/llm-planner.ts` | 规划器提示词（模型了解「有哪些操作」的唯一途径） |
| `tests/` | 测试与守卫；`tests/invoke-plugin.ts` 能在真实 cordis 上下文里跑工具 |
| `bundle/` | 发布产物。**`bundle/dist` 提交进仓库**，因为用户装的就是它 |
| `docs/modules.md` | 每个源文件做什么 |

## 动手前先知道的三件事

这三条来自这个项目真实踩过的坑，它们解释了为什么有些测试看起来「管得这么宽」。

**一、静默失败是主要的敌人。** 返回一个看起来合理的错值，比返回 `null` 危险得多——
`columnToNumber('Sheet1')` 曾返回 229493717，喂给 `getColumn` 直接堆爆进程。
转换函数要在入口校验，不要指望下游。

**二、同一份知识不要有两份手写实现。** 两份一定会漂。收敛到单一权威，再用测试断言
「每个消费者都必须覆盖清单里的每一项」。这个项目已经因为违反它出过三次事故。

**三、Windows 和 Linux 不一样。** CI 在两边都跑。换行、路径大小写、重命名一个打开中
的文件，行为都不同。

## 最常见的贡献：加一个 `excel_*` 工具或操作

一个能力被写在**多个**地方，漏掉任何一个都不会报错，只会有人按文档用的时候找不到。

1. **注册工具**：在 `src/tools/*.ts` 里 `name: 'excel_xxx'`。不要绕过
   `registration.ts` 的 `guardedContext` 包装——宿主拒绝一个工具时，会让这个插件注册过的
   **全部** effect 一起回滚，而且完全静默。
2. **如果是新的 `excel_operate` 操作**：先在 `src/operation-schema.ts` 加 schema
   （那是唯一权威），再在 `src/operations.ts` 的 `switch` 里加 `case`。
   **不要另建一份操作清单**——那是第三份副本。
3. **同步文档**（四处）：
   - `README.md` 的工具表
   - `docs/usage.md` 的工具表
   - `bundle/README.md` 的工具清单 —— **npm 页面上显示的就是它，最容易被忘**
   - 如果是操作，还要补 `src/llm-planner.ts` 的 `OPERATION_CATALOG`
4. **加测试**。守卫测试会替你检查上面这些，但不会替你检查行为对不对。

守卫测试清单：`docs-sync`（工具表 ↔ 已注册工具，双向）、`i18n-coverage`（每个 `t()` 都有
译文）、`readme-anchors`（页内链接指向存在的标题）、`docs-live-facts`（活文档不许写测试数）、
`plan-schema`（schema 声明的必填字段都真的被校验）。

## 测试

```sh
npm test                              # 全部
node --test tests/operations.test.ts  # 单个文件
TZ=UTC npm test                       # 提交前请跑一遍
```

**`TZ=UTC` 不是可选的。** 开发机常在 UTC+8、CI 跑 UTC，硬编码日期的断言会「本地绿、
CI 红」。写日期断言时，从读取端推导期望值，不要写字面量。

**测试要自足。** 不要依赖外部 CLI（Windows 默认没有 `unzip`）；要读 xlsx 内部 XML，
用 `fflate.unzipSync`。

**断言渲染后的结果，而不是实现里的字面量。** 否则测试会跟着实现一起漂，把错误行为钉住。
`tests/example-workbook.test.ts` 里那条钉住 `npm run demo` 输出的测试就是这么写的。

## 国际化

面向用户的文案用 `t('中文原文', { param })`，词条以中文原文为键；插值写 `{name}`。

两条不能破的规矩：

- **绝不在模块顶层调用 `t()`** —— 它会在 `setLanguage` 之前求值，把语言冻结在导入时刻。
  常量表存中文原文，查表时再翻。
- **只翻「消息」，不翻「数据」** —— 生成的工作表名、汇总标签、`_dsh_` 内部表名都有代码
  依赖它们，翻译会静默改变行为。用户自己的表名/列名/单元格内容永远原样保留。

## 提交与 PR

提交信息用 `type(scope): 一句话`，与现有历史保持一致：

```
fix(schema): enforce the required fields the schema already declares
feat(corpus): cover print titles and page breaks
docs: say when the demo was recorded, and what that means
```

- **一个 PR 做一件事。** 顺手重构请单独开一个。
- **改了 `src/` 就要重新构建并提交产物**：`npm run build:bundle`，然后提交
  `bundle/dist`。CI 会断言「提交的 bundle 与源码一致」，不一致就红——因为
  `bundle/dist` 是用户真正安装的东西，测试必须跑在它上面。
- CI 在 ubuntu 和 windows 上各跑一遍：类型检查 → 重建 bundle → 比对 `bundle/dist` →
  测试 → `npm pack --dry-run`。

## 不要写进仓库的东西

- **测试数量**。命令自己会打印，写下来只会过期（守卫测试 `docs-live-facts` 会拦）。
- **凭据**，任何形式，哪怕只是前缀。
- **维护者的事**：发版流程、npm 发布配置、本地开发环境的账号信息。仓库是对外的产品。

## 第一次贡献？

挑一个带 **`good first issue`** 标签的 issue，它们都带着明确的验收标准。不确定怎么下手就
先在 issue 里问，比做完再返工便宜。

## 行为准则

就一条：对事不对人。review 会直接说哪里不对、为什么，这是对代码的要求，不是对你。
