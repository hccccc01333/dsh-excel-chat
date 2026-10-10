# Changelog

## Unreleased

- **README 重构（中英文）**：照着高星仓库的做法改的（对比了 cline/cline、microsoft/markitdown、
  modelcontextprotocol/servers、awesome-mcp-servers 的结构）。

  | 之前 | 之后 |
  | --- | --- |
  | 25 个工具**平铺**成一张表，`excel_operate` 一格 **1500+ 字符** | 按**你想做什么**分四组；那格收进折叠块（默认隐藏） |
  | `Modules` 段 50 行开发者文件清单 | 移到 [docs/modules.md](docs/modules.md) / [docs/modules.en.md](docs/modules.en.md) |
  | 没有目录 | 加了目录，**14 + 10 个锚点全部校验过** |
  | 没有「为什么用它」 | 加了（照 markitdown 的 "Why Markdown?" 那段） |
  | 数字过时（423 项测试、100 个任务） | 483 → 489、100 → 113 |
  | 两个重复的安装段 | 合并成「安装与上手」 |

- **新增：README 的「最近更新」块由发版脚本生成**，不再手工维护。
  手写的摘要块在下一个版本发布时就过时了，而且没人会发现 —— 所以交给
  `scripts/readme-notes.mjs`，发版时从 CHANGELOG 的对应小节生成。

  **它只放版本号、日期和链接，不引用正文**：试过抽「前两条摘要」，结果不成立 ——
  有些条目第一行是完整摘要，有些写到一半才换行，抽出来是「…而语料任务的」这种半句；
  英文版还会引用中文（CHANGELOG 只有中文）。完整说明在 Release 页，GitHub 侧栏也会自动显示。

- **新增守卫 2 条**：
  - `readme-anchors.test.ts` —— 两份 README 的**每个页内链接都要指向存在的标题**。
    改标题是寻常编辑，没有别的东西会发现链接断了。实测把 `## 架构` 改名会立刻变红。
  - `readme-notes.test.ts` —— 生成器本身。第一版的正则要求标题**恰好等于 tag**，
    而实际标题带日期（`## v0.42.0 — 2026-10-10`）→ **匹配不到任何东西**，
    会安静地写出一个空块。

- 测试 483 → **489 通过**。

- **修复：GitHub Release 页面不写改了什么。** 发布 workflow 用的是
  `generate_release_notes: true`，而它生成的就一行 ——
  「Full Changelog: compare v0.41.0...v0.42.0」。于是**发布页对「改了什么」只字未提**，
  而专门为此写的 CHANGELOG 正文躺在仓库里，只有已经 clone 过的人才看得到。

  现在 release job 会**从 CHANGELOG.md 抽出该 tag 的那一节**作为 Release 正文
  （`body_path`），抽不到就**响亮地失败**，而不是默默退回自动生成的一行。

  **顺带修了一个我自己引入的顺序 bug**：新加的 `actions/checkout` 放在
  `download-artifact` 之后 —— checkout 默认会**清空工作区**，会把刚下载的
  `dist/*.tgz` 删掉，Release 就没附件了。已把 checkout 提到最前。
  （YAML 合法 ≠ 逻辑正确 —— 这个 bug 用 YAML 校验是查不出来的。）

- **README 增加指向 CHANGELOG 与 Releases 的指引**，中英文都加。

- 回溯更新了 **v0.42.0 的 Release 正文**：从 91 字符的自动占位符换成
  **2817 字符**的实际说明。

- 测试 483 通过。

## v0.42.0 — 2026-10-10

- **Benchmark：新增 7 条任务，语料 105 → 113 条，覆盖操作 31 → 39（共 77）。**
  补的都是「schema 里有、用户会用大白话问、但语料从没跑过」的：

  | 任务 | 断言要点 |
  | --- | --- |
  | `join-fill-from-lookup-sheet` | 两行都补齐（只处理首个命中会漏） |
  | `analysis-crosstab-live-sumifs` | 交叉表单元格是 `=SUMIFS(` 活公式 + 行列合计 |
  | `edit-unique-values-first-seen` | **首次出现顺序**（香蕉在梨之前，排序会调换） |
  | `edit-rank-column-live-formulas` | `=RANK(` 活公式 |
  | `edit-transpose-block` | 转置到旁边，**源数据不动** |
  | `edit-clear-range-keeps-neighbours` | 只清目标列，邻列必须保留 |
  | `edit-remove-empty-columns` | 空列被删、后续列左移 |
  | `format-copy-style-to-range` | 格式刷（加粗 + 填充同时断言） |

  期望值**全部先实测再写**，不是照文档猜的。

- **两条我写错的期望，被语料当场抓住**（这正是语料的价值）：

  1. `joinSheets` **不写输出列表头** —— 我假设它写，实际留空。已改为钉住实际行为，
     并注明「这处若变更是决定，不是漂移」。
  2. `removeEmptyColumns` 什么都没删 —— 因为我在任务里把中间列表头写成了 `备注`，
     那列**有内容**，本就不该删。我探测时用的是空表头，**探测夹具和最终任务不一致**。
     **教训：用最终要交付的那个夹具去探测。**

- `importCsv` / `exportCsv` **故意不做成语料任务**：它们指向工作簿之外的文件，而语料任务的
  operations 是静态的、夹具目录要到构建时才确定。这两个由单元测试覆盖。

- 测试 483 通过（新增 7 条任务，全绿）。

- **修复：37 条用户可见的警告文案根本没走 `t()`，中文用户看到的是英文。** 默认语言是
  **中文**，而这些文案是硬编码英文 —— 每一次改完文件后的操作提示、图表校验与公式体检的
  诊断，都是英文。

  **这个缺口上一轮我漏掉了**：当时量「本地化完整性」只数了 `t()` 的键，而**没包 `t()` 的
  字符串它根本看不见**，于是得出「只剩 1 个未翻译」的错误结论。

  现在 22 条操作警告 + 15 条诊断全部走 `t()`，中文原文为键、英文在词典里。
  含嵌套的「（还有 N 个）」也用 `t()` 处理，不把中文漏进英文句子。

- **新增守卫：警告文案必须经过 `t()`。** 检查的是**形状**而不是措辞：`message:` 的值里有
  长字符串字面量、又没调用 `t(`，就是有人会在错误的语言里读到它。实测能精确列出全部 37 处。

  **顺带发现：4 条测试把错误行为钉住了** —— 它们用正则断言英文文案，而中文是默认，
  等于把「中文模式下显示英文」写进了测试，所以它发现不了这个 bug。
  改为通过 `t()` 断言，两种语言下都成立，且钉住的是**用户真正会看到的**文案。

- 测试 482 → **483 通过**。

- **新增：把「四处同步」变成守卫测试。** 一个能力写在四个手维护的地方——`src/index.ts`
  注册工具、`README.md` 与 `docs/usage.md` 各有一张工具表、规划器提示词列出可用的操作。
  **它们之间没有任何连接**，所以「加了工具但忘了改某张表」要等到有人翻文档找不到才发现。

  现在四处都对着权威来源核对：工具表 ↔ 已注册工具（双向，既查漏也查「文档还在描述已删的工具」）、
  规划器提示词 ↔ schema 的操作集合。实测从 README 删掉一行会报 `excel_modify_chart`、
  把提示词里的 `rankColumn` 改名会报 `rankColumn`。

  **实现上的坑**（记下来，因为我犯了两次）：提取工具名**不能只匹配行首** ✗——
  表格里有合并行 `` | `excel_create_chart` / `excel_modify_chart` | ``，
  按行首匹配会把已记录的工具报成缺失 ✗。改为**从整行里抓所有反引号名** ✓。
  同理，规划器目录是**散文**（`insertRows/deleteRows/…`），不能按 `"op":"x"` 匹配 ✗。

- 测试 473 → **477 通过**。

- **重构（第一步）：把 `operations.ts` 的类型抽成 `operation-types.ts`。**
  `operations.ts` 长到 3273 行，把类型、dispatcher、59 个 handler、文件入口塞在一起。
  类型是**所有东西都要导入**的那部分，所以先搬它：`operations.ts` **原样再导出**
  （`export * from`），**没有任何调用方需要改**。3273 → 2976 行。

- **新增守卫：dispatcher 的 `case` 集合必须与 schema 的操作集合完全一致。**
  这两份名单是同一个东西的两份手写副本：schema 决定「调用方能发什么」，`switch` 决定
  「发了会发生什么」。**schema 有、case 没有 = 校验通过然后静默什么都不做**；
  **case 有、schema 没有 = 谁都到不了的死代码**。两者都不会自己变成测试失败——
  所以这条测试把两份名单都读出来对比（含重复 case 检查）。
  实测把 `case 'set'` 改名会精确报出 `unhandled: 'set'`。

  实现上有个坑值得记：`case` 不能按「任意缩进」匹配 ✗——handler 内部还有自己的 switch
  （criteria 操作符 `eq`/`gt`/…），会被误当成操作名 ✗。改为**限定在 dispatcher 函数体内、
  只取最浅缩进的 case**（嵌套的必然更深）✓。

- 测试 472 → **473 通过**。

## v0.41.0 — 2026-10-09

- **新增守卫：重规划反馈必须具体到「哪条断言、期望什么、实际什么」。** 「未达成」对规划器
  毫无可用信息，下一轮只能靠猜——那正是失败分类里 replan 那一条描述的
  （「第一轮失败后没有纠正」）。验证器的失败详情本来就带单元格、期望值和实际值，
  现在有测试把这三样都钉进下一轮收到的结论里；**把详情从结论里去掉会让它变红**。

- 测试 471 → **472 通过**。

- **修复：规划器提示词里的示例自己就是非法的。** 5 个示例把占位符写进了枚举位置——
  `"keep":"first|last"`、`"mode":"value|forward|left"`、`"role":"ops|product|data"`、
  `"direction":"asc|desc"`、`"mode":"contents|formats|all"`。**模型照抄这个形状就会产出
  非法值**，然后被 schema 当参数错误拒掉——提示词在自己制造它本该防止的失败。

  改为「给一个具体合法值 + 用文字说明其他选项」：形状合法，信息不丢
  （如 `"keep":"first"` 后附「keep 二选一：first 保留首个 / last 保留末个」）。

  新增守卫：**把提示词里每个 JSON 示例抽出来跑一遍 schema**，逐项核对枚举值、布尔类型、
  必填字段。实测把 `first|last` 放回去会立刻报
  `dedupeRows: keep="first|last" is not one of first/last`。

- 测试 470 → **471 通过**。

- **新增：枚举值与布尔的确定性归一。** `function: "SUM"`、`function: "求和"`、
  `keep: "FIRST"`、`bold: "true"`、`direction: "降序"` 和 schema 里的常量是同一个东西，
  只是写法不同——以前一律当参数错误拒掉。

  归一规则**从 schema 派生**，不另抄一份：模块加载时把 schema 里所有 `enum` 与
  `type: 'boolean'` 的**路径**走一遍（含 `style.hAlign`、`metrics[].function`、
  `keys[].direction`、`rules[].type` 这类嵌套），再按同一批路径回写。
  这样 schema 加了枚举，归一会自动跟上，**不会出现「schema 加了、salvage 忘了同步」**。

  匹配顺序：先大小写不敏感精确匹配（`SUM`/`First`/`DataBar`），再查**一张小而一对一**的
  中文别名表（`求和`/`平均`/`计数`/`居中`/`升序`/`运营`…）。

  **两条自我约束**：① 已经是合法值的一律不动；② **别名表里每个目标都必须是该字段的合法值**——
  有专门的不变式测试拿 schema 校验它，所以写错一个别名会变成测试失败，而不是静默改写。
  不在允许集里的值（如 `中位数`）**不猜**，原样留给 schema 去报错。

- 测试 465 → **470 通过**。新增 4 条归一测试 + 1 条别名不变式测试；
  **停掉归一会让其中 3 条变红**。

- **新增：表头名当作列字母时的确定性 salvage。** 规划器看到的是 profile 里的**表头**
  （`区域`、`金额`），所以很自然会写 `groupColumn: "区域"`，而 schema 要的是 `"B"`。
  这本来会以 `invalid column letter: 区域` 失败——**意图毫无歧义，答案就在已经拿到的 profile 里**。

  现在按表头解析成列字母（`column` / `groupColumn` / `valueColumn` / `outputColumn` /
  `scoreColumn` / `rowColumn` / `columnColumn` / `metricColumn` / `sourceKey` / `targetKey` /
  `lookupKey` / `columns[]`，以及 `metrics[]` / `summaryColumns[]` / `keys[]` / `criteria[]`
  里的 `column` 和 crosstab 的 `metric.column`）。

  **两条自我约束**：① 已经是列字母的一律不动；② **表头在同一张表里出现两次就拒绝解析**
  （名字不再唯一标识一列，猜一个就是在用户文件里写一个看起来合理的错答案）。找不到表头时
  也原样保留，让失败和以前一样响。

  接在规划器路径上——那里 profile 是现成的，**零额外开销**。工具调用路径保持原样（它们
  本来就明确报错）。`tests/plan-schema.test.ts` 加 5 例；**停掉解析逻辑会让其中 3 例变红**。

- 测试 460 → **465 通过**。

- **修复：17 个任务的断言写死了任务描述里从没提过的输出表名。** 例如
  `preset 产品模板：报表+色阶。` 却断言一张叫 `订单-产品分析` 的表。**规划器把活干对了、
  只是把表叫了别的名字，就会被判失败**——成功率因此在惩罚「没猜中一个没说的名字」，
  而那不是一个值得衡量的能力。v0.37 记录里那批「良性命名差」的 Argument 失败主要就是这个。

  处理：**把表名写进描述**（用户在意时本来就会说「输出到汇总表」），让任务成为完整规范。
  实测「断言写死表名但描述没点名」的任务 **17 → 0**，语料仍 100/100。
  新增守卫测试：任何断言引用了计划新建的表，描述里就必须出现这个名字。

- **修复：没有断言的任务会被真空判成满分。** `success` 的判据是
  `checksPassed === checksTotal`，而 `checksTotal === 0` 时 `0 === 0` 成立 → 白拿一个成功；
  精度聚合里也有同样的 `checksTotal === 0 ? 1`。改为要求 `checksTotal > 0`。
  当前语料没有这种任务，但没有任何东西阻止它发生——补了一条守卫：**每个任务都必须有断言，
  每条断言都必须至少声明一个可检查条件**，这样语料里的笔误会变成测试失败，而不是更好看的数字。

- 测试 458 → **460 通过**。

- **修复：`excel_undo` 对格式类编辑是个报告成功的空操作。** 回滚原本只有一条路——
  把 `.patch.json` 里的 `oldValue`/`newValue` 对调、当成 `set` 重放。而那份日志只记录
  **单元格值差异**，所以：

  | 编辑 | patch 日志 | 旧回滚的效果 |
  |---|---|---|
  | `style`（加粗/填充） | **0 条**（值没变） | **什么都没做，却报告成功** |
  | `insertRows` / `deleteRows` | 值差异 | 值还原了，**行列结构留在原地** |

  根因是**回滚的粒度错了**：每次编辑都经 ExcelJS 重写整个包，一次操作改变的东西
  （格式、合并、行列结构、工作表顺序、透视表锚点）**远多于**值差异能看到的。

  改为**按文件快照恢复**——`writeWorkbookSafely` 在覆盖前已经存了 `<path>.bak`，
  那就是编辑前的完整字节。`excel_undo` 优先用它，**还原单元格内容、公式、格式、合并、
  行列结构与透视表**；没有快照时（该次编辑是写出新文件而非覆盖）才退回按值回滚，
  并在结果里**明说它只还原了内容**。工具描述同步改写。

- **修复：就地编辑的审计日志一直是空的。** `operateWorkbookFile` 先写文件、再读
  「修改前」状态——`outputPath === path` 时两次读的都是**已被改写的文件**，于是
  `diffCellMaps` 恒为空，`excel_undo` 的兜底重放对它永远无效。改为**先读后写**。

- 测试 452 → **458 通过**。新增 `tests/undo.test.ts`（6 例），其中一条专门钉住
  「按值回滚救不回格式」这个限制，避免兜底路径被误当成完整回滚。

- **修复：编辑会静默销毁透视表。** ExcelJS 不建模透视表，重写时 `xl/pivotTables/*`
  与 `xl/pivotCache/*` 连同所有引用一起消失。现在这些部件会**逐字节复制**回输出，
  并把三处引用（sheet 的 pivotTable 关系、`workbook.xml` 的 `<pivotCaches>`、
  workbook rels 的 pivotCacheDefinition 关系）与 `[Content_Types].xml` 的覆盖补回。

  **注入的关系一律重新分配 r:id**：原文件的 `rId1` 在 ExcelJS 刚写完的文件里毫无意义，
  沿用它会静默指向 ExcelJS 恰好放在那里的东西。实测原文件用 `rId3`、输出用 `rId6`
  （因为 ExcelJS 占了 rId1–rId5），引用同步改写。

  **本机用 Excel 实测确认**：编辑后的文件 `Workbooks.Open` 无修复提示，
  `PivotTables().Count` 为 1——透视表真的还在。

- **新增：写入前的结构校验 + 拒绝写入。** 合并要交叉修改四个文件的引用，改错就是
  产生 Excel 打不开的包——比丢透视表更糟。所以合并后先校验
  （每个 rels 的 Target 都能解析到实际部件、每个 `r:id` 都有声明、
  Content_Types 覆盖有对应部件、**有 pivot 部件就必须能追溯到 workbook**），
  **不过就抛错拒绝写入**。配合原子写，原文件完好。
  校验是结构级的、不需要 Excel，所以 CI 能跑。

  *这条护栏当场生效过一次*：校验器自身有个 `_rels/.rels` 的路径解析 bug，
  于是第一次端到端跑就被拦下——没有写出损坏文件。

- **新增：原子写 + 覆盖前备份。** `writeFile` 先截断再写，崩溃/磁盘满/被杀会在原位置
  留下截断的 xlsx。改为**同目录临时文件 + rename 覆盖**（同目录才同文件系统，
  rename 才是原子的）；失败清临时文件并抛错，**目标保持原样**。覆盖已存在文件时
  自动备份到 `<path>.bak`。接入 `operations.ts` / `patch.ts` / `health-report.ts`。

- **丢失预警不再包含透视表**（现在会保留），改为只报告真正会丢的：
  切片器、时间线、VBA、宏表、customXml、表单控件、嵌入对象、SmartArt。

- 测试 446 → **452 通过**。新增 `tests/preserve.test.ts`（含用真实 Excel 生成的
  透视表夹具做端到端，CI 无需 Excel）与 `tests/safe-write.test.ts`。

- **修复：`report` 的汇总表把总额算成 3 倍。** `report` 会先往源表插小计行，再把
  **含小计行的整块区间**交给 `aggregateReport`；而后者把分组列里**每个不同的值**都当成
  一个分组——小计行的标签（`华东 汇总`）和总计行的标签（`总计`）都是「不同的值」，
  于是每个都变成幽灵分组，最后那行总计再把真实分组**重算一遍**：

  | 行 | 内容 | 值 |
  |---|---|---|
  | 2 | 华东 | 14246 |
  | 3 | 华东 汇总 | 14246 ← 与上行完全重复 |
  | … | （华北/华南/西南 同样各重复一次） | |
  | 10 | 总计 | `SUMIFS(…, "总计")` → 54324 ✓ |
  | 11 | 总计 | `=SUM(B2:B10)` → **162972 = 真实总额的 3 倍 ✗** |

  改为在枚举分组时跳过小计行，识别依据是 **`SUBTOTAL(` 公式**而不是标签文案——
  标签是工作簿数据、且是中文，按文案匹配既脆弱又会挡住 i18n。
  用 Excel 自己算出的结果核对过：修后为 4 个真实区域 + 1 个总计，总计 = 54324 =
  独立计算的真实总额。

  新增 `tests/report-summary.test.ts`（3 例）。**撤掉修复会让其中 2 例变红**，
  不是空过。这是「岗位报表自动化」主打的产出物，总额错 3 倍会直接毁掉信任。

- **新增：100 任务评测的公开结果表，别人可以自己跑一行。** 数字以前是手写在文档里的
  markdown 表——会漂，别人也没法产出可比的一行。现在：

  - `bench-results/*.json` 是**原始跑分产物**（就是 `node tests/invoke-llm-benchmark.ts`
    的 stdout），`docs/benchmark-results.md` 由 `npm run bench:report` **生成**，
    `bench-results/table.csv` 供机器读。
  - **加一行 = 放一个 JSON 文件**，不需要改代码；文档里给了 DeepSeek / BAI / 本地
    Ollama 三种跑法（Ollama 零 API 成本、数据不出网）。
  - `npm run bench:corpus` 是**不调用模型**的免费入口（100/100 语料回归，约 5 秒），
    先确认语料本身是好的再花钱跑。
  - **两条护栏，都来自踩过的坑**：① 没配 key 时 runner 不会报错退出，而是把 100 个任务
    全记成 `execution` 失败——那看起来就是「模型得 0 分」，所以生成表时**拒绝**
    「全 execution + 0% 成功率」的结果；② 缺 model / date 的行直接报错，
    不让它变成表里的空白格。
  - `tests/bench-report.test.ts` 守住「提交的表与产物一致」——**手改表会让它变红**。
  - 文档里也写清了直接跑**不要加 `--test`**（会往 stdout 混 TAP 行，输出就不是纯 JSON）。

- **宣传素材改为生成器产出**（`npm run assets`）。图里的每个数字都来自**真跑一遍引擎**
  （`collectEvidence()` 真的执行 `autofixWorkbookFile`），`tests/assets.test.ts` 把图钉在
  测量值上——**故意改错期望值会让它变红**。主张改为「Excel 的活，说一句就干完」，
  公式体检从主打降为「所以可信」。

- **修掉一个自己引入的 CLI 缺陷**：为让脚本可被测试导入而加的入口守卫，在 Windows 上
  永远为假（`fileURLToPath` 给绝对反斜杠路径，`process.argv[1]` 是相对路径），
  于是 CLI **静默不执行**——而测试全绿，因为测试走的是导入路径。
  改用 `pathToFileURL(process.argv[1]).href` 比较。

- 测试 429 → **439 通过**。

## v0.40.0 — 2026-10-08

- **新增：用户可见文案支持中英文切换**，走 dsh 配置项而不是环境变量或跟随消息语言。
  在 profile 的 `cordis.patch.yml` 里覆盖即可：

  ```yaml
  - id: vera
    config:
      language: en        # zh（默认）| en
  ```

  取值沿用宿主自己的词汇（`@deepseek-ai/dsh-client-locale` 用的就是 `zh` / `en`），
  不配置即 `zh`，行为与之前完全一致。

  实现放在 `src/i18n.ts`，**词条以中文原文为键**（`t('中文原文', { param })`）而不是自造
  key：调用点保持可读，而且没翻译的会自动回退中文，不会在界面上漏出 `health.anomalies`
  这类 key。共 268 条词条、191 个调用点，覆盖体检报告、数据洞察、能力菜单、修复摘要、
  公式解释、验证断言、规划器提示、失败归因、doctor 自检，以及三个斜杠命令的描述与提示。

  **刻意不翻译三类东西**，因为它们不是「给人看的消息」：

  - **工作簿里的数据产物**：汇总行标签（`总计`/`合计`/`平均`）、生成的工作表名
    （`-汇总`/`-筛选`/`-交叉表`）、岗位模板名。它们写进文件，**而且有代码依赖它们**——
    `src/patterns.ts` 正是用 `/^(总计|小计)$/` 识别并跳过汇总行，`src/health-report.ts`
    的隐藏表名 `_dsh_体检报告` 由 `validator` 按 `_dsh_` 前缀跳过。翻译这些会让行为
    静默改变，也会让同一个文件在不同语言下产生不同结构。
  - **规划器提示词**（`src/llm-planner.ts`）与**工具/参数描述**：给模型的输入，
    few-shot 例子是调过的，而且模型两种语言都读。
  - **你自己的表名、列名、单元格内容**：永远原样保留。

  两个实现细节值得记：**列表分隔符也随语言走**（中文 `、`/`；`，英文 `,`/`;`），
  否则英文句子里会出现 `A、B、C` 和 `SUMIFS（…）`；**绝不在模块顶层调 `t()`**——
  `explain.ts` 的函数描述是模块级常量表，顶层求值会把语言冻结在导入时刻，
  所以改成查表时再翻。

- **修复：失败归因不再依赖中文报错文案。** `src/failure-taxonomy.ts` 原来用正则匹配
  规划器抛出的**中文**文案（`operations 数组`、`op 字段`、`缺少必填`、`必须是数字`）
  来判断失败类型。这在上一条的翻译面前是个真障碍：**把文案翻成英文，规划类错误就会
  静默变成执行类错误**，归因结果不再可信。

  没有回避翻译，而是拆掉耦合：`plan-schema.ts` 改抛带 `kind` 的 `PlanSchemaError`，
  分类器按**类型**判断，只对引擎自身的英文消息保留文本兜底。
  连带修了一处：`agent.ts` 会把错误重新包装成普通 `Error`，**类型会在链路上丢失**，
  所以加了 `wrapRoundError` 保留它——否则类型化只做对了一半。

  **通用教训：凡是「拿人类可读文案当判据」的地方，翻译都会踩雷；先换成结构化信号再翻。**

- **文档：新增英文版并加语言切换。** `README.md` / `bundle/README.md`（npm 页面）各配一份
  `README.en.md`，顶部放「简体中文 | English」互链；`bundle/package.json` 的 `files`
  加入英文版，否则 npm 页面上那个链接指向的文件不在包里。写英文版时顺带修掉根 README
  三处过期事实：`operation-schema.ts` 的「27 操作」→ **77**、`index.ts` 的
  「exposing eight tools」→ **25 tools**、版本锁定示例 `@0.23.0` → `@0.39.9`；
  「Run tests」原来列了 18 个具体文件（过期子集），改成跑全量。

- **文档：重写 npm 页面介绍。** 原来的介绍比 GitHub 首页薄得多，且同一段安装命令出现了
  两遍。现在明确两个差异化点（**编辑后自动体检**，而不是改完就交差；**核心能力纯 XML 层、
  跨平台、不依赖本机 Excel**），补上快捷命令表与 `dsh-excel-chat-doctor` 自检入口、
  可靠性数据与文档索引；`description` 重写到 255 字符内（npm 搜索会截断）。

- 测试 423 → **429 通过**。新增 `tests/i18n.test.ts`，其中**第一条守卫测试扫描全部
  `src/*.ts` 的 `t('…')` 调用点并断言每个都有译文**——加新消息而忘记翻译时它会在 CI 前
  就报出来（实测精确抓到过一处遗漏）。`tests/load-bundle.test.ts` 增加一条走
  **构建产物 + 真实 cordis 上下文**的用例，验证配置接线本身（而不只是词条表）没坏。

## v0.39.9 — 2026-10-08

- **修复：goal 模式下只改格式/元数据的步骤一律被判为「未达成」。** agent 循环用
  `workbookFingerprint` 判断「这一步到底改没改文件」，而它是**单元格值 + 粗体 /
  数字格式 / 填充**的快照。可绝大多数操作改的东西**根本不在单元格网格里**：

  | 操作 | 旧指纹 |
  | --- | --- |
  | addComment 加批注 | 无变化 |
  | setColumnWidth 设列宽 | 无变化 |
  | freezePanes 冻结窗格 | 无变化 |
  | autoFilter 加筛选 | 无变化 |
  | conditionalFormatting 条件格式 | 无变化 |
  | dataValidation 数据校验 | 无变化 |
  | addTable 建表 | 无变化 |
  | setZoom 设缩放 | 无变化 |
  | setHyperlink 更新已有链接 | 无变化 |

  实测 10 个操作里 **9 个检测不到**。而循环里有
  `if (!changed || …) verdict = { achieved: false }`——**于是这些步骤被强制判失败**，
  即使操作成功、文件完全正确。goal 模式下「冻结首行」「加个筛选」「给表头加批注」
  这类请求因此必然失败，还会白白消耗重规划轮次。

  改为**按部件内容取哈希**（解压 xlsx，对排序后的部件做 sha256）。这是**按构造完备**
  的——新操作只要写出任何东西就会被覆盖——而枚举式快照必须手工维护、必然漂移，
  正是这个项目反复出现的缺陷类型。

  比较前两侧都先过一遍 exceljs：它写出时会重新序列化，所以「Excel 写的文件」与
  「exceljs 写的文件」即使语义相同也逐字节不同，不归一化就会恒判「有变化」。
  实测 exceljs 输出**逐字节确定**（同一输入两次写出一致，`dcterms:created` /
  `modified` 被原样保留而非重新打时间戳），所以归一化是安全的。空操作现在正确
  判为「无变化」，把单元格设成它已有的值同样如此。

- **修复：`excel_profile` 把日期列的样本值显示成本地语言英文串。** 样本来自
  `String(date)`，于是出现 `Thu Jan 15 2026 09:30:00 GMT+0800 (中国标准时间)`
  ——依赖机器语言环境、与工具其他部分的 ISO 表示不一致，换台机器还会变成
  `Eastern Daylight Time`。现在统一为 ISO。

- **修复：`excel_undo` 回滚会把数字、布尔、日期都变成文本。** 补丁日志里存的是
  `cellContent` 的产物（字符串），而写入端 `applyPatchesToWorkbook` 直接
  `cell.value = patch.newValue`，于是撤销一次编辑会把 `42` 恢复成**字符串** `"42"`
  ——数字不再能求和，日期变成文本，超链接只剩显示文本。

  **这个 bug 特别阴**：`cellContent` 对数字和「数字文本」的渲染完全一样，所以既有的
  往返测试（比较 `readWorkbookCells` 的字符串）**看不出任何差异**。只有断言单元格
  *类型*才能抓到它。

  现在写入端复用与 `set` / `importCsv` 相同的推断：`contentToCellValue` 从
  `operations.ts` 移到 `cellContent` 旁边，让序列化与反序列化成对，三个写入方共享
  同一份规则。实测七种单元格形状（整数 / 小数 / 布尔 / 日期 / 公式 / 文本 / 超链接）
  撤销后类型全部保住。

  已知限制：日志只存字符串，所以「内容看起来像数字的文本单元格」会被还原成数字——
  这与 `set` / `importCsv` 的行为一致。

- **修复：`exportCsv` 把四种单元格形状写成垃圾。** 序列化用的是 `String(raw)`，
  而 exceljs 里日期、超链接、富文本、错误值都是对象：

  | 单元格 | 旧输出 |
  | --- | --- |
  | 超链接 `{text:'官网', hyperlink:…}` | `[object Object]` |
  | 富文本 `{richText:[…]}` | `[object Object]` |
  | 错误值 `{error:'#REF!'}` | `[object Object]` |
  | 日期 | `Thu Jan 15 2026 17:30:00 GMT+0800 (中国标准时间)` |

  日期那一条尤其糟：它是**依赖机器时区与语言环境**的英文串，既不可解析也不可移植。
  现在统一走 `cellContent`（日期→ISO、超链接→显示文本、富文本→纯文本），错误值再经
  `plainContent` 取回裸 token（Excel 导出 CSV 就是这么写的）。

- **修复：日期读出来之后写不回去。** `cellContent` 产出 `toISOString()`，所以
  `excel_read` 把日期显示成 `2026-01-15T01:30:00.000Z`；但写入端的正则
  `^(\d{4})-(\d{2})-(\d{2})([T ](\d{2}):(\d{2})(:(\d{2}))?)?$` **不接受毫秒与 `Z`**，
  于是把读出来的值原样写回会**降级成文本**——而且因为文本恰好长得一样，这个损坏
  几乎看不出来。现在正则接受完整 ISO 8601（含小数秒与区标识）：带区标识时按绝对
  时刻解析，不带时按本地墙钟解析（与 exceljs 的读写约定一致，实测往返精确一致）。
  顺带修好毫秒被截断的问题。

- **修复：公式注入防护是单向的，导出再导入会污染单元格。** `exportCsv` 给看起来像
  公式的文本加 `'` 前缀，但 `importCsv` 不去掉，于是文本 `=1+1` 导出成 `'=1+1`、
  再导入就成了带撇号的 `'=1+1`——以 `+` 开头的手机号这类文本尤其常见。现在加了
  `unguardFormulaInjection` 作为防护的精确逆操作，并且**把还原后的值强制写成文本**：
  交回给 `writeContent` 会重新推断出公式，正好把防护要挡的东西复活。

- **修复：`excel_task` 的步骤缺字段时抛裸 TypeError，而不是指出缺了哪个字段。**
  `excel_operate` 的 `operations` 用的是 `excelOperationSchema`，宿主会在调用前校验，
  畸形操作到不了 handler。但 `excel_task` 的 `steps[].operations` 声明是
  `{ type: 'object', additionalProperties: true }`——**没有任何 schema 保护**，
  于是漏字段的步骤一路走到 handler 内部崩掉：

  ```
  TypeError: Cannot read properties of undefined (reading 'toUpperCase')
  ```

  这条信息既没说是哪个操作，也没说缺哪个字段。现在两个入口都先过一遍
  `sanitizeOperations`——**复用规划器已有的那套校验与修复**（`sanitizePlan`），
  而不是再写一份：同一条规则有两份实现必然漂移。同样的输入现在得到
  `sortRange 的 keys[0].column 缺失`。

  顺带补上一个可用性缺口：`excel_operate` 收到不带表名的 `range`（如 `"A1:B1"`）
  时会报 `range requires a sheet`，尽管这在 schema 里是合法的（`range` 只是个
  字符串）。现在会按第一个工作表补全前缀，与规划器路径一致。

  动手前已核对四张 `REQUIRED_*` 表要求的每个字段在 schema 里都标了 `required`，
  所以这层不会拒绝任何 schema 允许的输入。

- **修复：编辑后体检会把「提到错误值的普通内容」误判成错误单元格。**
  `detectErrorValues` 用的是**不锚定**的正则，所以错误 token 出现在内容的任何位置
  都算命中，置信度还是 1。实测这些都会被误报：

  | 单元格内容 | 旧行为 |
  | --- | --- |
  | `备注：#REF! 已修复` | 报错 |
  | `=IF(A1="#N/A","missing",A1)` | 报错 |
  | `=IFERROR(A1/B1,"#N/A")` | 报错 |

  最后一条最说明问题：**这个公式存在的意义就是处理 #N/A，却被判成含错误的单元格**。
  它还让 `validate` 与 `excel_find_errors` 对同一份文件结论不一致——
  `findErrorCells` 读的是类型化的 `cell.value.error`，从不误报。方向正好与上一版
  相反：那次是体检漏报 `#N/A`/`#NAME?`，这次是体检过报。

  现在只认真正会出现的几种形状：整值就是错误 token（允许 `=#REF!` 这种公式体）、
  `readWorkbookCells` 产出的 JSON 信封 `{"error":"#REF!"}`、以及以 `=` 开头且
  **剥掉字符串字面量后**以错误 token 结尾的公式结果文本。最后一条的「剥掉字面量」
  是关键：`=IFERROR(A1/B1,"#N/A")` 剥掉后剩 `=IFERROR(A1/B1, )`，不再以 token 结尾。

- **修复：模型把列名当列字母传入时会把进程打崩（OOM）。** `columnToNumber`
  对任何字符串都不做校验，`columnToNumber('Sheet1')` 会算出 **229493717**
  （S=19,H=8,E=5,E=5,T=20,1=−15 的进位结果）。`groupColumns` 只检查
  `from < 1 || to < from`，这个值两个条件都通过，接着
  `sheet.getColumn(229493717)` 让 exceljs 分配巨型数组——**Node 直接以 4GB
  堆上限崩溃**。也就是说，一个把列名写进列字母字段的模型输出就能终结用户的会话。
  （旁证：`numberToColumn(229493717)` 回绕出 6 个字母的 `SHEESK`，
  之前那些 `Invalid column letter: SHEESK` 的乱码报错都是这条链的下游产物。）

  现在 `columnToNumber` 要求 `/^[A-Z]{1,3}$/` 且结果不超过 16384（XFD），
  否则抛出 `invalid column letter` / `column out of range`。它被 60 多处调用，
  其中不少直接接收模型传入的列名（`groupColumn`、`targetKey`、`lookupKey`、
  `from`/`to`…），所以在这一处收紧即可全线受益。

- **修复：`parseCellId` 接受行号 0。** Excel 行号从 1 开始、到 1048576 结束，
  但 `A0` 会被解析成 row 0，超界的大行号也照收——它们最终同样走到
  `sheet.getCell` 的越界分配。现在行号必须落在 1..1048576。

- **修复：CSV 导入会给每张表多写一整行空行。** `parseCsv` 把**末尾换行**当成了
  新记录的开始，而真实 CSV 文件都以换行结尾——所以一个 3 行的 CSV 导入后
  `sheet.rowCount` 是 4，第 4 行是 `A4=''`。更根本的是**往返不幂等**：
  `stringifyCsv` 永远追加 `\r\n`，于是 `parseCsv(stringifyCsv(x))` 必然多一行，
  导出再导入会不断长行（`exportCsv` 无 range 时正是用 `sheet.rowCount` 定范围）。
  现在只在 `field !== '' || row.length > 0` 时才补最后一行，`stringifyCsv([])`
  返回空串（原来返回 `\r\n`）。端到端复验：3 行 CSV（末尾带换行）导入后
  `rowCount = 3`。

- **修复：自引用公式查不出循环引用。** `buildDependencyGraph` 的 `addEdge` 里
  有一句 `if (from === to) return`，把自环丢掉了——于是 `A1 = A1+1`
  （Excel 最常见的循环引用，通常是少打一行的笔误）**`edges` 和 `cycles` 都是空的**，
  `validator` 的 `circular-reference` 异常永不触发；`A1 = SUM(A1:A5)` 同理。
  自引用是合法输入，不该在构图时被优化掉。删掉那句早退后 `findCycles` 天然处理
  自环，`traceDependencies` 靠 `seen` 集合也不会死循环。

- **修复：`parseRangeRef` 会把不带表名的区间静默解析成错值。** 原来 `!` 是可选的，
  贪婪的表名分组会吃掉整个区间：`A1:B2` 被解析成**表名 `"A1:"`、只覆盖 B2 一格**。
  返回一个看似合理的错值比返回 `null` 危险得多——调用方会继续拿它去校验单元格。
  改为要求 `!` 后 `A1:B2` 返回 `null`，图表校验会据此报 `invalid-range`，
  与 `A1` 原本就返回 `null` 的行为也一致了。

- **修复：公式注入防护漏掉 `\t` 与 `\r` 前缀。** 按 OWASP CSV Injection 的清单，
  Excel 会先剥掉前导制表符/回车再判断是否公式，所以这两个也是攻击前缀；
  而 `stringifyCsv` 不会给含制表符的字段加引号，`\t=cmd` 会原样落进 CSV。

- **修正语料中两处循环引用夹具**（`formula-repair-multi-sheet`、
  `formula-repair-cross-sheet-offset`）。它们把公式从 `margin()` 的 4 列布局
  （A=产品/B=收入/C=成本/D=毛利，`=B2-C2` 在那里是正确的）搬进 3 列布局
  （A=收入/B=成本/C=毛利）却忘了把引用左移一列，于是「毛利」写成 `=B2-C2`——
  **公式在 C2 里引用 C2**，语义上算的是「成本 − 毛利」。这是夹具自身的缺陷，
  此前被上面那句丢弃自环的代码掩盖了；修好构图后两个任务的完整性立刻从 1 掉到
  0.98，才把它暴露出来。按语义左移一列修正，原有的偏移错误与断言意图保持不变。

- **新增 5 个测试文件，并给 8 个既有文件补了用例，共 +73 例**，补上此前没有直接测试
  的模块：`tests/csv.test.ts`（16 例，含端到端 `importCsv` 回归、四种单元格形状的
  导出、导出导入往返）、`tests/graph.test.ts`（14 例）、`tests/charts.test.ts`（8 例）、
  `tests/formula.test.ts`（10 例，列字母与单元格 id 的边界校验）、
  `tests/workbook-fingerprint.test.ts`（12 例：九个「不在网格里」的操作都必须被指纹
  检测到，空操作必须判为无变化——这条守卫防止指纹再次漂移）；
  `tests/operations.test.ts` 的 OOM 回归与日期读写回归、
  `tests/diff.test.ts` 的 `rolling a patch back restores the cell types, not just
  the text`（断言单元格*类型*，既有往返测试比较字符串所以抓不到）、
  `tests/profile.test.ts` 的日期样本回归、
  `tests/error-values.test.ts` 的 3 例误报回归、
  `tests/plan-schema.test.ts` 与 `tests/task.test.ts` 的入口校验回归，以及
  `tests/load-bundle.test.ts` 的 `excel_task names the missing field instead of
  crashing on a malformed step`（**跑构建产物 + 真实 cordis 上下文**，
  验证的是可达路径）和 `tests/file-benchmark.test.ts` 的语料守卫
  `no corpus fixture ships a circular formula`——遍历 100 个夹具断言无环，
  这条守卫当初就能抓到那两处夹具缺陷。测试 350 → **423 通过**。

  日期相关断言一律写成**与时区无关**的形式（从读取端推导期望值，而不是硬编码
  ISO 串）：CI 跑在 UTC 上，而开发机在 UTC+8，硬编码的断言会「本地绿、CI 红」。
  本地可用 `TZ=UTC node --test tests/*.test.ts` 复现 CI 的时区条件。

## v0.39.8 — 2026-10-07

- **修复：工作表已有扩展列表时，sparkline 与批注会被插到错误的位置。**
  exceljs 每加一个数据条或色阶，就会在**条件格式规则内部**写一个嵌套的 `<extLst>`
  ——而 `preset` / `report` 正是加这些的。原实现用
  `xml.replace('</extLst>', …)` 追加 sparkline 扩展，替换的是**第一个**闭合标签，
  于是扩展被塞进了 cfRule 内部：**Excel 不会渲染它，sparkline 等于静默失效**。
  `patchSheetForComments` 有同样的缺陷——`xml.replace(/<extLst/, …)` 把
  `<legacyDrawing/>` 插到第一个 `<extLst` 之前，会让工作表**不符合 schema**。

  改为按标签深度定位「直接挂在 `<worksheet>` 下的那个 `<extLst>`」；
  `legacyDrawing` 则插在 `tableParts` 与 `extLst` 中**靠前者**之前
  （schema 要求它在两者之前）。走真实操作链路复验（加数据条 → 加 sparkline）：
  扩展已落在条件格式之后，文件仍可被读回。

- **新增两个直接测试文件**，补上此前只有间接覆盖的模块：
  - `tests/xml-postprocess.test.ts`（14 例）：批注四件套（comments / VML / rels /
    Content-Types）齐全、XML 转义、VML 的零基行列、showFormulas 的两种情形、
    sparkline 逐行配对与行列不匹配的报错、**嵌套 extLst 的插入位置**，
    以及重开文件确认可用。内含一个零依赖的 XML 标签平衡校验器——这个模块的 XML
    全是手写的，格式良好性必须被测到。
  - `tests/patterns.test.ts` 扩到 14 例：`detectPatternAnomalies` 的偏移异常与
    expected/actual 偏移、少数派槽位不算「缺失」、汇总行与 SUBTOTAL 行豁免、
    绝对引用、按列独立分析。

  测试 325 → **347 通过**。

- **加固 `annotateWorkbookXml` 的属性解析**：它读的是**别人写的文件**，原来用
  `/<Relationship[^>]*Id="…"[^>]*Target="…"/` 这种**位置敏感**的正则，一旦写入方把
  `Target` 放在 `Id` 之前就会静默失配（定位不到工作表 → 批注/sparkline 静默不注入）。
  改为逐个属性独立查找（`attributeOf`），同时支持单双引号。已用真实重排过的 rels
  验证：属性顺序颠倒后仍能正确定位。

- **加固 `withShowFormulas` 的三种边界**，原来只处理「第一个 `<sheetView>`」：
  - **空 `<sheetViews></sheetViews>`**：原实现会再插一个 `<sheetViews>` 块 →
    **产生两个列表，XML 非法**。现在把 view 插进已存在的那个列表里。
  - **一个表有多个 `<sheetView>`**：原来只补第一个，其余 view 仍显示计算结果。
    现在每个缺该属性的 view 都补上，已有该属性的不重复补。
  - 顺带修正：`showFormulas` 是**每个 view 各自**的属性，不是表级属性。

  测试 347 → **350 通过**。

## v0.39.7 — 2026-10-07

- **修复：编辑后的公式体检漏掉 `#N/A` 与 `#NAME?`。** `patterns.ts` 里硬编码了一份
  自己的错误值正则 `/#(?:REF|DIV\/0|VALUE|NAME\?|N\/A|NULL|NUM)!/g`——它**要求结尾
  必须有 `!`**，而 `#N/A` 结尾是 `A`、`#NAME?` 结尾是 `?`，两个都匹配不上；
  `#GETTING_DATA` 更是完全不在候选里。后果：**VLOOKUP/XLOOKUP 查找失败**与
  **函数名拼错**这两类最常见的错误，`validator` 的编辑后体检会**放行**，而同一个
  工作簿用 `excel_find_errors` 却能查出来——两个工具对同一份文件给出不同结论，
  恰好落在插件「自动发现静默错误」这个承诺上。

  根因是同一份知识存了两遍然后漂了。现收敛到单一来源 `src/error-values.ts`：
  `audit.ts` 从它 re-export（对外接口不变），`patterns.ts` 从它派生正则
  （值会做转义，避免 `?` 被当成量词）。新增 `tests/error-values.test.ts` 6 例，
  核心一条是**「清单里的每个值都必须被编辑后体检检出」**——以后加值忘了同步消费者
  会直接测试失败。

- **错误值清单补全到 17 个**。上一条把清单收敛成唯一来源后，实测发现它**本身就是短的**：
  `findErrorCells` 从不查这份清单（它接受文件里任何错误 token），所以它一直能报出
  `#SPILL!`；而**派生自清单的编辑后体检报不出来**——两个工具再次对不上。逐个核实后补齐
  Excel 365 动态数组与链接数据类型（`#SPILL!` `#CALC!` `#FIELD!` `#BLOCKED!`
  `#UNKNOWN!` `#CONNECT!` `#BUSY!`）与 Python in Excel（`#PYTHON!` `#TIMEOUT!`）。
  顺带修正了 `excel_find_errors` 的工具描述——它此前列出的错误值比实际能查的少。

- **公式列里的「文本型千分位数值」现在能被判为硬编码断裂**。实测确认：正常打
  `1,000` 会被 Excel 存成数字、读出来是 `"1000"`，判定本来就对；漏的只是**以文本形式
  存储**的数值（从网页粘贴、CSV 导入），它们保留千分位分隔符因而绕过了数值判定。
  这恰恰是「数字存成了文本」这个 Excel 常见问题。判定放宽为接受 `\d{1,3}(,\d{3})+`，
  同时**畸形分组（`1,00`、`1,0000`、`12,34`）仍被排除**。此前这条行为既无文档也无测试，
  属疏忽而非设计。

  测试 325 通过（新增 7 例：`tests/patterns.test.ts` 6 例 + 清单覆盖 1 例）。

## v0.39.6 — 2026-10-07

本次发布包含 v0.39.5 的全部改动（dsh 0.2.0 静默跳过 bundle 的根因修复、更宽的
peer 区间、doctor 的 `bundle-compatibility` 检查、`DSH_EXCEL_CHAT_STATUS` 状态文件），
外加下面这些：

- **发布流程改为 npm trusted publishing（OIDC）**。`publish.yml` 不再使用
  `NODE_AUTH_TOKEN`：由 GitHub 为该工作流签发短期凭证，仓库里不再存放长期凭据，
  发布也不再需要人工过 2FA。`--provenance` 也不再显式传入——公开仓库的 OIDC 发布
  会自动生成来源证明。
- **新增一键发版脚本** `scripts/release.mjs`：

  ```sh
  node scripts/release.mjs X.Y.Z     # 或 npm run release -- X.Y.Z
  ```

  改 `bundle/package.json` 与根 `package.json` 的版本 → 把本段定版为
  `## vX.Y.Z — 日期` → 重建 `bundle/dist` → 跑全量测试 → 提交 → 打 tag → 推送，
  之后由 CI 完成测试、构建、发布与 GitHub Release。`--dry-run` 只校验不改动，
  `--no-push` 只本地提交。脚本会拒绝脏工作区、空的 `## Unreleased` 段、
  已存在的 tag，以及不在 `master` 上的发布。
- 测试 312 通过。

<!-- 在这里写本次改动；`node scripts/release.mjs X.Y.Z` 会把它定版并打 tag。 -->

## v0.39.5 — 2026-10-06

### 找到并修复了 #5 的根因：dsh 0.2.0 会**静默跳过** peer 不匹配的 bundle

`@deepseek-ai/dsh-app-boot@0.2.0-rc.2` 的 `loadProfileDirectory` 里，每个 bundle 都要过
一次兼容性检查：

```js
const issue = evaluatePluginCompatibility(bundleManifest, exemptions);
if (issue !== undefined && !issue.exempted) throw new Error(pluginCompatibilityWarning(issue));
...
} catch (error) {
    skippedBundles.push({ packageName, reason: String(error) });   // ← 吞掉，什么都不打印
}
```

`evaluatePluginCompatibility` 会把本包 `peerDependencies` 里所有
`@deepseek-ai/dsh` / `@deepseek-ai/dsh-*` 与**运行时 dsh 版本**
（`getDshRuntimeVersion()`，即 app-boot 自己的版本）做
`semver.satisfies(runtime, range, { includePrerelease: true })` 比对。
**不匹配 → bundle 被丢进 `skippedBundles` → patch layer 不应用 → 行不挂载 →
没有任何工具、也没有任何报错。** 文档原文：*"nothing is printed"*。

**为什么 A/B 看起来那么干净**：豁免是按 `name@version` **精确匹配**存的
（profile 的兼容性文件）。报告者当年为 `dsh-excel-chat@0.38.1` 授过豁免，
所以 0.38.1 照常加载；0.39.0 是**新版本号**，没有豁免 → 被静默跳过。
`0.38.1` 和 `0.39.0` 的 peer 区间**完全相同**（都是 `^0.1.0-rc.6`），
所以这从来不是「0.39.0 引入了 bug」，而是「0.39.0 没被豁免」。

实测判定（复刻宿主的检查逻辑，runtime = 0.2.0-rc.2）：

| 版本 | peer 区间 | 结果 |
| --- | --- | --- |
| 0.38.1 / 0.39.0 | `^0.1.0-rc.6` | **不兼容 → 静默跳过**（0.38.1 靠豁免才活着） |
| 0.39.4 | `^0.1.0-rc.6 \|\| ^0.2.0-rc.1` | 兼容 |
| **0.39.5** | **`>=0.1.0-rc.5 <0.3.0`** | 兼容，且覆盖 0.1.x / 0.2.x 全线 |

- **`peerDependencies` 的 `@deepseek-ai/dsh-*` 改为 `>=0.1.0-rc.5 <0.3.0`**，
  覆盖 0.1.0-rc.5 … 0.2.1-alpha.1（含预发布），同时不冒充支持尚未验证的 0.3.x。
- **新增回归测试 `tests/compatibility.test.ts`**：复刻宿主的检查规则，
  对已知内核矩阵断言区间必须通过，并**反向断言 `^0.1.0-rc.6` 会被拒绝**——
  确保这个测试不会空过。
- **`dsh-excel-chat-doctor` 增加 `bundle-compatibility` 检查**：直接调用宿主导出的
  `evaluatePluginCompatibility` / `getDshRuntimeVersion` / `readProfileCompatibility`，
  提前告诉你「dsh 会不会静默跳过本插件」，而不是等装完发现工具没了。
  内核早于 0.2.0 时该检查不存在，会明确说明而不是误报。

### 其余诊断增强（同版本）

- `dsh-excel-chat-doctor` 增加 `tool-registration` 自检：用记录型假宿主跑一遍真实
  `apply()`，报出注册了几个工具与每个失败注册的带标签错误。
- 新增 `DSH_EXCEL_CHAT_STATUS` 环境变量：设置后 `apply()` 结束时把注册结果写到该路径
  （`{ applied, tools, failures, at }`），用来区分「`apply` 从没被调用」与
  「`apply` 调了但注册失败」。不设置则完全不碰磁盘。
- 测试 308 → 312 通过。

## v0.39.4 — 2026-10-06

> 发布过程绕了一段路，记录在此。直接 `npm publish` 连续失败，**每个版本号都是
> 首次 PUT 就报** `409 Cannot publish over previously staged version`；换 tag、
> 绕代理、直连、从 `bundle/` 发都一样，registry 无故障。根因是 `~/.npmrc` 用的是
> **bypass-2FA 的 granular access token**，而 npm 已把这类 token 的
> 发布能力收缩为**只能 staged publish、需维护者用 2FA 批准**（npm changelog
> 2026-07-31 *Restricting npm bypass-2FA granular access tokens*）——该账号已提前
> 生效；0.39.0 能发出去只是因为它赶在切换之前。
>
> 关键误解：那几次报 409 的 PUT **其实都已经成功 stage 了**（`npm stage list`
> 当时还没显示出来，约 1 分钟后才转 `staged`）。所以
> **0.39.1 / 0.39.2 / 0.39.3 / 0.39.4 四个版本最终都发布了，代码完全相同**
> （均为 136 文件，只差 `package.json` 里的版本号），`latest` 指向 0.39.4。
> 它们都是这份修复，没有中间态。
>
> **以后发版直接用 staged publishing**（`npm stage publish` 不需要 2FA）：
>
> ```sh
> cd bundle
> npx --yes npm@11 stage publish                 # 不需要 2FA
> npx --yes npm@11 stage list dsh-excel-chat     # 等 status 变成 staged
> npx --yes npm@11 stage approve <stage-id>      # 维护者完成 2FA
> ```
>
> 也可以直接在 npmjs.com 的 **Staged Packages** 标签页点 Approve。
> **注意：一次报错的 `npm publish` 也会留下 staged 条目**，所以别重复发——
> 直接 `stage publish` 一次，然后去批准，不要靠重试版本号。

- **注册路径不再静默失败**（回应 #5：0.39.0 在 DSH 0.2.0-rc.2 上所有 `excel_*`
  工具消失）。此前每个工具都是
  `ctx.effect(() => ctx.tools.register(defineTool({...})), 'tool:xxx')`，而 cordis
  对 `effect` 回调里的抛错**不做隔离**：`composeError` → `handleError` 原样重抛，
  错误冒出 `ctx.effect()`、冒出 `apply()`，插件随之失败并**回滚它注册过的全部
  effect**。因此任何一个工具在宿主侧被拒绝，都会把其余 24 个一起带走，用户看到
  的正是「插件像没加载过」——没有工具，也没有任何报错。现在每个注册各自隔离：
  失败的那一个被跳过、其余照常注册，并打出带标签的
  `tool:xxx failed to register: <stack>`。
- 加载过程**留下可查的记录**：`plugin loaded` 与结束时的
  `ready: N tools registered` 同时写宿主 logger 与 console；有失败时打
  `PARTIAL: N tools registered, M failed -> <标签>`。不再存在「看起来一切正常」
  和「什么都没有」这两种无法区分的状态。
- 整个注册路径外包了一层 try/catch：非 effect 路径的失败也会带标签记录后再抛。
- `peerDependencies` 的 `@deepseek-ai/dsh-*` 由 `^0.1.0-rc.6` 放宽为
  `^0.1.0-rc.6 || ^0.2.0-rc.1`。原区间**不覆盖当前 DSH Desktop 2.0.17 内核的
  0.2.0-rc.2**，会让包管理器为插件再装一份 0.1.x 的 `dsh-tools`。已逐个核对
  0.2.0-rc.2 与 0.1.0-rc.6 的 `defineTool`、`parameterSchemaSpecToJsonSchema`、
  `valueSchemaSpecToJsonSchema`、`assertSupportedJsonSchema` 完全一致，且 25 个
  工具的 schema 在 rc.2 下全部通过注册校验（`register()` 会调
  `assertSupportedJsonSchema(output.schema)`）。
- 新增 `tests/registration.test.ts`（5 例）：一个注册失败不影响后续注册、
  抛错不会逃出 `ctx.effect`、失败同时上报 logger 与 console、摘要区分完整与
  部分加载，以及**用真实 `apply()` 验证宿主拒绝一个工具时其余 24 个仍然起来**。
  测试 305 通过。

## v0.39.0 — 2026-10-05

- 文档补齐：`excel_operate` 的 **CSV 导入/导出**此前在 README 与使用指南里完全没提
  （v0.28.0 就实现了），已补；`bundle/README.md`（npm 页面显示的那份）只列了 12/25
  个工具，连 `excel_task`（goal 闭环）和 `excel_autofix` 都不在，已补全为 25 个并按
  用途分组，与根 README 对齐。
- 规划器操作目录补齐 **10 个此前对模型不可见的操作**：`clear`、`definedName`、
  `importCsv`/`exportCsv`、`mailMerge`、`setZoom`/`showGridLines`/`showFormulas`、
  `protectSheet`/`unprotectSheet`。它们早已实现（邮件合并、工作表保护、命名区域
  在 README / 使用指南里都有场景），但从未进过规划器提示词，模型**规划不出来**
  ——「用模板给每笔订单生成发货通知」「导出成 CSV」「保护工作表」「缩放 150%」
  全都会落空。目录与参数速查已同步补齐。
- 新增**规划器目录覆盖守卫测试**：从 `excelOperationSchema` 推导全部操作名，
  断言每一个都出现在规划器实际发出的 prompt 里。以后新增操作忘了同步目录会
  直接测试失败，而不是静默变成模型够不到的能力。测试 300 通过。
- `excel_operate` 的 `insertImage` 只给单边尺寸时**按原图比例推出另一边**。原先是把
  缺失的一边直接设成 100px，代码注释还声称这样「不会悄悄拉伸比例」——实际恰恰相反：
  给一张 100×200 的图传 `width: 300`，会渲染成 300×100 的变形图。改为从图片头读出
  原始尺寸后按比例换算（比例读不出来且只给了一边时才报错）。工具 schema 的尺寸说明
  同步更新。测试 299 通过（新增 2 例：只给宽、只给高）。

- 文档同步：`docs/usage.md` 的工具表补上遗漏的 `excel_trace`、`excel_find_errors`、
  `excel_validate_charts_visual`、`excel_export_pdf`（此前只有 README 有），
  `excel_operate` 操作清单补插入图片 / 显示公式视图 / 按颜色与自定义序列排序 /
  固定宽度分列，平台说明补 PDF 导出与视觉评审；`docs/benchmark.md` 说明 LLM 基准
  走标准 `/chat/completions`，任何 OpenAI 兼容端点都能接（含本地 Ollama，
  数据不出网）。
- 测试不再依赖外部 `unzip`：`tests/operations.test.ts` 里 4 个断言原始 XML 的
  用例原先用 `execSync('unzip -o ...')` 解包 xlsx。`unzip` 在 Windows 上默认
  不存在，受限环境（无 `cmd.exe` spawn 权限的沙箱）也不允许起 shell，这些用例
  在作者机器之外必挂。改为用已在依赖里的 `fflate` 直接读 zip 条目（新增
  `readZipEntry` helper），测试自足、跨平台。测试 297 通过。
- `excel_operate` 的 `sortRange` 支持**高级排序模式**：
  - `by: 'fill' | 'font'` + `color`：把带指定颜色的行排到最前（缺 `color` 会明确报错）；
  - `customList: ["高","中","低"]`：自定义序列排序，列表外的值排在后面。
  - 实现上把颜色分组与序列名次**在收集阶段就折算成可比较的标量**，
    因此原有比较函数一行未改，也没有为这两种模式引入特例分支。
  - 顺带修了一个既有缺陷：**排序原先只搬运值、不搬格式**。Excel 排序是整行移动，
    而按颜色排序恰恰依赖格式跟着走（否则值移走了、颜色留在原地，排序毫无意义）。
    现在样式与值一起搬。
  - 测试 297 通过（新增 3 例：自定义序列、按颜色且格式跟随、缺 color 报错）。
- 新增 `excel_find_errors` 工具：列出所有**错误值**单元格，附产生它的公式与按错误码的计数，
  可按工作表限定；能区分「真错误值」与「文本恰好长成错误样」，后者不报以免变成噪音。
  - ⚠️ 关键实现细节：**带公式的错误单元格 `type` 是 `Formula` 而不是 `Error`**，
    错误藏在 `value.result` 里 —— 只判断 `type === Error` 会漏掉所有由公式产生的错误，
    而那恰恰是最值得报的一类。改为读 value 形状，两种情况都覆盖，也不依赖类型枚举。
  - 不复用 `readWorkbookCells`：它会把值统一字符串化，无法区分真错误和文本。
  - 测试 294 通过（新增 `tests/audit.test.ts` 3 例）。
- `excel_operate` 新增 `showFormulas`：让工作表保存为「显示公式」视图（`show:false` 切回）。
  - exceljs 的 `SheetViewXform` 只渲染固定属性白名单，**没有 `showFormulas`**，
    给 `view.showFormulas` 赋值会被静默丢弃；且 exceljs 仅在已有视图设置时才写
    `<sheetViews>`，普通工作表里根本没这个元素。
  - 因此改由 `xml-postprocess` 在保存后注入：已有 `<sheetView>` 就补属性，
    没有就按 schema 顺序在 `<dimension>` 之后插入整个 `<sheetViews>` 块。
  - 修了一个连带缺陷：`annotateWorkbookXml` 的调用点自己写了一遍「有没有注解」的判断
    （只查 comments/sparklines），新增第三类注解后必须同步，否则注入根本不触发。
  - 测试 291 通过（新增 2 例：注入生效且公式不变、show:false 不注入）。
- 新增 `excel_trace` 工具：追踪单元格的公式依赖链路，`direction` 可选 `precedents`（引用，
  它读了谁）/ `dependents`（从属，谁读了它）/ `both`，`depth` 控制层数（默认 1，即 Excel
  单击一次的行为），每格返回当前值与深度，并附带循环引用。
  - 这是 Excel「追踪引用/追踪从属」的**数据版**：那两个功能画的是箭头，而箭头是界面状态、
    **不写进 xlsx**，所以只能以链路数据的形式复现，反而更适合给模型和脚本消费。
  - 复用既有 `buildDependencyGraph`（体检内核用的同一个依赖图），新增 `traceDependencies`
    做 BFS 遍历并处理环，`truncated` 标记提示链尾之外还有内容。
  - 图内部的 key 是 canonical 大写形式，工具层映射回工作簿的真实写法后再返回。
  - 测试 289 通过（新增 `tests/trace.test.ts` 5 例）。
- `excel_operate` 的 `splitColumn` 支持**固定宽度分列**：用 `widths: [6, 8, 4]` 代替
  `delimiter`（两者互斥，都不给或都给会明确报错）；超出最后一个宽度的文本作为追加列
  保留而不丢弃。规划器参数示例与 README 已同步。测试 284 通过（新增 3 例）。
- `excel_operate` 新增 `insertImage`：在指定单元格嵌入 png/jpeg/gif 图片。
  - 纯 XML 层实现（exceljs 负责写 `xl/media/` 与 drawing 部件），**跨平台、不依赖本机 Excel**；
  - 可指定像素宽高；省略时从图片头读取原始尺寸（PNG / GIF / JPEG），避免拉伸变形；
  - 已有图片会保留（在加载出的工作簿上追加，而非重建）；单张上限 20MB；
  - 不支持格式、`file`/`base64` 同时给或都不给时明确报错。
- 同步更新：规划器操作目录与参数示例、`plan-schema` 必填字段校验、`excel_operate`
  工具描述、README 操作清单。测试 281 通过（新增 3 例：嵌入与可读性、原有图片保留、参数校验）。

## v0.38.1 — 2026-09-27

- 修复 issue #4：`excel_operate` 的 `set` 传入**原生标量**（数字/布尔/日期）时
  抛 `content.trim is not a function`。工具 schema 对 `cells` 声明
  `additionalProperties: true` 并在描述里承诺「numbers/dates/booleans are
  typed」，但 `writeContent` 无条件调用 `.trim()`，标量在类型推断前就崩了；
  走 planner 的路径因为 `sanitizePlan` 会 `String()` 化而幸免，只有直接调用
  该工具的路径（含遍历数值列的调用者）会踩到。
  - `writeContent` 对非字符串内容直接赋值，保留标量类型；字符串仍走
    trim + 类型推断；
  - `set` 的 `cells` 类型由 `Record<string, string>` 放宽为
    `Record<string, CellContent>`（新增导出 `CellContent = string | number |
    boolean | Date | null`），与 schema 契约对齐，避免调用方在 TS 层被误导；
  - 边界降级而非抛错：`null` 清空单元格；`NaN`/`Infinity` 降级为文本（Excel
    无法表示，直接写入会让工作簿打不开）；其他对象类型降为文本；
  - 回归测试 2 例：类型标量写入（含公式、日期、小数）与 null 清空 /
    非常规标量降级后文件仍可读；已确认回退该修复时这 2 例会失败。
- 测试 278 通过（`node --test tests/*.test.ts`），`tsc -p bundle/tsconfig.json`
  类型检查通过。

## v0.38.0 — 2026-09-07

- **Verifier 2.0：规划器机器可查断言**（goal 模式验证合取的第 2 层确定性
  防线，继 v0.37.0 清零 Verification 误判后进一步压 Argument/Planning 类
  失败）：
  - `AgentPlanner.plan` 可返回 `{steps, assertions}`：assertions 是
    `{"id":"汇总!B2","startsWith":"=SUMIFS("}` / `{"id":"订单!A1","expect":"区域"}`
    形式的机器断言，复用 `verifyWorkbookAssertions` 对执行后的文件确定性
    校验（值/公式前缀/填充/加粗/数字格式/对齐/换行）；
  - 合取判定：LLM 判定 ∧ 确定性校验 ∧ 规划器自身断言——断言失败会阻止
    "已达成"并携带失败明细回喂重规划；断言通过不能推翻负判定（不引入新的
    误判方向）；
  - `sanitizeAssertions` 净化层：补工作表前缀、数字/布尔 expect 归一化为
    序列化文本、丢弃缺 id/无可检查字段/类型不支持的条目，容错不炸循环；
  - 规划器提示词要求每轮输出 3-8 条断言（只断言计划产出的内容）。
- 修复 issue #3（大表在右侧面板显示不全）：
  - 行数显示上限 200 → 2000、新增列数上限 100，超出时显示截断提示而非静默丢数据；
  - 内联视口 420px → 560px，新增「全屏」按钮（fixed 覆盖层占满窗口，Esc /
    按钮退出，窗口缩放自动重建网格）；
  - 修复 AA+ 宽表编辑列名错位（原 `String.fromCharCode(65 + colIndex)` 在
    第 27 列起生成非法引用，改为 Excel 进制的列名换算）。
- 主动审计修复（自检发现的数据/性能问题）：
  - `transpose` 就地转置（target 落在 source 上/与其重叠）读到被覆盖的单元格
    导致数据损坏——改为先快照整个源区域再写出；非方阵就地转置先清源块避免
    残留；补 self-transpose 回归测试；
  - `copyRange`（含 valuesOnly）复制到与其重叠的目标区同样损坏——改为快照源块；
    补重叠 copy 回归测试；
  - `hideRows` / `groupRows(collapse)` 传巨大 to/end（如 200000）会物化海量
    空行致文件暴涨——钳制到已用行范围并给出 warning；补钳制回归测试；
  - `freezeFormulas` / `copyRange valuesOnly` 对无缓存结果的公式（插件刚写出的
    文件常见）不再抹成空：freezeFormulas 跳过、copyRange 回退为复制公式；
  - 手写进公式的表名统一 Excel 引号限定（`qualifySheetName`）：含空格/中文/
    特殊字符的表名不再让 RANK/SUMIFS/COUNTIFS/aggregateReport 公式失效；
  - `renameSheet` 改裸表名用词边界正则，不再把 `AA!` 误伤成 `ZA!`（重命名
    "A"→"Z"）；补回归测试；
  - `uniqueValues` 去重键类型感知：数字 1 / 文本 "1" / 布尔不再混为一类，
    无缓存结果的公式按其文本区分而非全部塌成空值；
  - `crosstab`：count/counta 不再把 metric 区域塞进 COUNTIFS（奇数个参数会致
    每格 #VALUE!）；行/列 key 用原始值写回，日期/数字维度可正确匹配；
  - `setHyperlink` 内部跳转改用 `HYPERLINK("#'表'!A1","文本")` 公式（ExcelJS 对
    `#` 内部链接会同时写 location 与 External rel 致跳转失效）；
  - `addSparklines` 以解析后的真实表名归档并按行对齐到 location 行（原按数据
    行），带引号/改名的表不再在保存期崩溃；
  - 批注 VML 的 [Content_Types] 类型修正为 `...vmlDrawing`；多作者批注不再全部
    归到首个作者（补 authors 列表 + 每注 authorId 映射）。
  - 面板截断提示由 ref 改为 state（首帧渲染读不到 ref 的时序问题）。
- 测试规模 262 → 268（agent 合取 4 用例 + sanitizeAssertions 2 用例 +
  提示词/透传断言）。

## v0.37.2 — 2026-09-07

- 基准运行器加固（真实 100 任务全量跑通的前提）：
  - `runLlmBenchmark` 支持 `outFile` JSONL 断点续跑（逐任务落盘、重跑跳过
    已完成、崩溃任务自动重试）与 `onProgress`（stderr 进度）；
  - 瞬态错误重试（429/500/503/fetch failed，默认 4 次线性退避）+ 任务间隔
    节流（`LLM_BENCH_RETRIES` / `LLM_BENCH_RETRY_DELAY` / `LLM_BENCH_TASK_DELAY`）；
  - 逐任务单跑脚本 `scripts/bench-run-missing.sh`，托管端点限流时只补缺失。
- LLM 基准新增 BAI 供应商（api.b.ai，OpenAI 兼容）：`LLM_PROVIDER=bai` +
  `BAI_MODEL`（glm-5.3-flash / qwen3.8-flash），`BAI_MAX_TOKENS` 默认 4000
  适配思考型模型。
- **glm-5.3-flash 全量 100 任务：成功率 86%（DeepSeek 基线 52%，+34pp）、
  准确率 88.5%、完整性 99%；分析类 28% → 84%，公式 68% → 95.5%，工作流
  39% → 77.8%；Verification 误判清零**（证据标准 + 按表轮询快照采样的
  直接效果）。数字为提示词 + 模型双变量叠加，见 docs/benchmark.md 归因。
- 全量实测暴露并修复的规划容错：
  - `style.horizontal/vertical` 别名到 `hAlign/vAlign`（sanitizePlan 确定性
    salvage）；
  - 口语化表名（"订单表"）匹配到精确表名（"订单"）——renameSheet/
    dedupeRows 等的 sheet/oldName/name 字段；
  - `fillSeries` 的 `range` 别名到 `target`；
  - 提示词补 preset.role 必填、outputSheet 默认命名、fuzzyMatch/fill/
    fillSeries/renameSheet 参数速查（全量跑缺示例的操作恰是缺字段重灾区）。
- 测试规模 258 → 262。

## v0.37.1 — 2026-09-05

- 修复 issue #2 两个插件生命周期问题（报告基于 v0.34.1 发布产物）：
  - 插件标识对齐：`export const name` 由 `vera-formula-validator` 改为
    `dsh-excel-chat`（与 bundle 包名、cordis.patch.yml 的 name 一致），
    加载日志同步更新；
  - 生命周期绑定：全部 27 处注册（21 个 excel_* 工具 + 3 个命令 +
    system prompt 段落 + 2 处间接注册）改用 `ctx.effect(() => register(...))`
    包裹，disposer 交给插件 fiber 收集——插件停用/更新时工具自动注销，
    不再残留；load-bundle 新增「卸载后工具消失」回归测试。
- 测试规模 256 → 258。

## v0.37.0 — 2026-09-04（本地版本，未发布）

- LLM 规划/验证质量（基准失败归因：Verification 33/52、分析类参数复杂）：
  - 规划器提示词全面升级：操作目录补齐 22+ 新操作并按用途分组，参数速查
    覆盖 joinSheets/crosstab/rankColumn 等，新增 4 个分析类完整 few-shot
    示例（汇总/两表关联/交叉表/排名+版式）与分析任务规则（禁用 set 写死
    汇总、source 必须含表头与全部数据行、crosstab metric 是对象等）；
  - 验证器提示词按目标类型给证据标准：汇总/透视必须看到输出表头+分组行+
    SUMIFS 活公式；样式/隐藏/冻结等快照看不到的要求必须核对计划中确有对应
    操作；新增"set 写静态数字不算汇总"反例；
  - `cellSnapshot` 改为按工作表轮询采样（96 格），新建输出表的证据不再被
    第一个表挤掉——这是 Verification 误判的最大杠杆；
  - `sanitizePlan`/`normalizeOperation` 认识全部新操作：crosstab 扁平
    metric 字段自动合并、joinSheets 标量数组自动包装、必填字段缺失明确
    报错回喂重规划。
- 能力盲区补齐（`excel_operate` 新增 4 操作 + 新模块 `src/xml-postprocess.ts`）：
  - `addComment`（单元格批注）：ExcelJS 只读不写，保存后用 fflate 后处理
    zip 注入 comments XML + VML 形状 + rels/ContentTypes/legacyDrawing，
    Excel 正常渲染、ExcelJS 读回 `cell.note`；
  - `addSparklines`（每行趋势迷你图）：手写 x14 sparklineGroups 扩展 XML
    （line/column/stacked、高低点颜色），dataRange 与 locationRange 按行
    一一配对；
  - `rowPageBreaks` / `clearPageBreaks`（打印分页符）：ExcelJS 只写不读
    rowBreaks（brk id 为 0-based，断言走 XML）。
- 测试规模 249 → 253。

## v0.36.0 — 2026-09-01

- `excel_operate` 再增 10 个操作（跨平台）：
  - 编辑：`copyStyle`（格式刷：复制字体/填充/边框/对齐/数字格式到区域）、
    `freezeFormulas`（公式转值：就地替换为缓存结果）、`uniqueValues`
    （提取列去重值，首现顺序）、`rankColumn`（活 RANK 公式排名列）、
    `unmergeAll`（一键取消全表合并）；
  - 视图：`setZoom`（缩放 10-400，保留冻结窗格）、`showGridLines`
    （网格线开关）；
  - 打印：`headerFooter`（&-code 页眉页脚，支持奇偶页/首页不同）；
  - 工作簿：`moveSheet`（按 orderNo 重排工作表标签）、`setWorkbookProperties`
    （作者/标题/关键词 + `recalcOnOpen` 打开时强制重算，写入 calcPr）。
- 新工具 `excel_export_pdf`：用本机 Excel COM 把整个工作簿或单个工作表导出
  PDF（Windows；只读打开，不改动源文件）。
- 引擎发现：exceljs 的 `worksheets` getter 按 `orderNo` 排序——重排只需改
  orderNo，不碰私有 `_sheets`；`headerFooter` 是 worksheet 顶层属性而非
  pageSetup 子字段；`calcPr fullCalcOnLoad` 会写 XML 但读取时不解析（测试
  直接断言 XML）。
- 测试规模 233 → 243。

## v0.35.0 — 2026-08-25（本地版本，未发布）

- `excel_operate` 新增 12 个操作，补齐常用 Excel 能力：
  - 行列可见性与分组：`hideRows` / `hideColumns`（隐藏/取消隐藏，空行也可靠
    持久化）、`groupRows` / `groupColumns`（大纲分级 1-7，可折叠，自动同步
    sheet 的 outlineLevelRow/Col 属性让 +/- 控件正常显示）；
  - 布局：`autoFitColumnWidths`（按内容自适应列宽，CJK 计双宽、min/max 可调）、
    `unfreezePanes`（取消冻结）；
  - 编辑增强：`transpose`（转置粘贴，公式相对引用随新位置偏移）、`copyRange`
    新增 `valuesOnly`（选择性粘贴-数值，取公式缓存结果）、`clearRange`
    （contents / formats / all 三档，对应 Excel「清除」菜单）；
  - 分析：`joinSheets`（两表精确匹配多列回填，VLOOKUP 无公式版，
    key 归一化 trim+lowercase，支持 missValue 兜底）、`crosstab`（二维交叉
    透视表，单元格用活 SUMIFS/COUNTIFS 公式引用源区域与输出表头地址，
    sum/count/counta 自动加总计行列，average 包 IFERROR 防 DIV/0）；
  - 其他：`setHyperlink`（外部 URL 与站内 location 跳转链接）、`printTitles`
    （打印时每页重复标题行/标题列）。
- `style` 操作扩展：删除线 strikeThrough、文字旋转 textRotation、缩小字体
  填充 shrinkToFit、缩进 indent；边框样式枚举扩展 hair / mediumDashed /
  dashDot / mediumDashDot / dashDotDot / slantDashDot。
- 引擎适配说明：ExcelJS 空行仅设 hidden 不落盘——hideRows/groupRows 对无值行
  补写默认行高保证序列化；批注（comment）ExcelJS 只读不写，暂不提供。
- 测试规模 174 → 233（operations.test.ts 新增 14 个覆盖保存往返的用例）。

## v0.34.0 — 2026-08-15（本地版本，未发布）

- 按评审 P0 转向 evaluation-driven：新增文件级真实任务语料 ExcelBench lite
  （100 个职场场景：编辑 35 / 分析 25 / 公式 22 / 工作流 18）+ `runFileBenchmark`
  运行器（任务成功率 / 平均准确率 / 完整性率 / 修复数），
  `tests/invoke-file-benchmark.ts` 可打印报告。
- 语料暴露并修复 3 个真 bug：
  - `parseFormula` 不识别中文表名跨表引用，依赖图把 `订单!B2:B7` 当同表
    引用导致误报循环引用——扩大未加引号表名匹配；
  - 汇总/总计行被列 pattern 误报为结构异常——新增汇总行守卫；
  - subtotal 汇总列中的数据行被误报 hardcode——含 SUBTOTAL 的列跳过
    硬编码判定。
- 新增中文跨表引用回归测试；测试规模 171 → 174。
- `excel_task` 新增 goal 模式（Agent 闭环）：LLM 规划操作步骤 → 逐步执行
  （公式体检+修复）→ LLM 验证器判断目标是否达成 → 未达成带上一轮结果
  重新规划，最多 `maxRounds` 轮；100 任务语料可作为该闭环的评测基准。
- 真实 LLM 规划基准（deepseek-chat，100 任务）：任务成功率 50%、完整性率
  85%（编辑 66% / 公式 64% / 工作流 39% / 分析 24%）；修复 2 个引擎问题
  （autofix 无修复时 final 文件不存在、验证器缺执行后快照）并强化规划器
  提示与容错（sheet 前缀、字段别名、cells 转字符串）；完整数字与失败
  归因见 [docs/benchmark.md](docs/benchmark.md)。
- 增强版：plan schema 校验/自动修复层（`sanitizePlan`，缺失必填字段明确报错
  并回喂重规划）、确定性验证并入闭环（公式异常 + 值/样式指纹变化与 LLM 判断
  合取）、分析类提示词强化、基准 maxRounds=3；重跑 100 任务：成功率 48%、
  完整性 89%、分析类 24% → 36%，引擎异常噪音清零。README 新增架构图。
- 包装：用真实 DeepSeek Harness Web + 真实模型（DeepSeek-V4-Flash）录制功能
  演示 GIF（`assets/demo.gif`）：一次对话完成“经营报表 + 公式体检修复 +
  汇总重建”，挂到 README 顶部。
- 表格展示：新增 `excel_preview`，把指定表/区域渲染为 Markdown 表格（对话内
  直接可见）+ HTML 预览文件；跨平台、无新依赖。
- 右侧 Web 面板（M1）：插件自带 client 模块（`dsh.client` +
  `exports["./client"]` → `/plugins/dsh-excel-chat/client.js`），注册
  `tool.call.toolview` 渲染器，点开 excel_* 工具行在右侧 details 列显示表格；
  已在 harness web 实测：boot 图加载、模块注册并 materialize。方案与验证
  见 [docs/web-panel.md](docs/web-panel.md)。
- 右侧 Web 面板 M2/M3：表格单元格双击编辑 → `inputActions.setDraft` +
  `submit` 提交给 agent 执行 `excel_operate.set` 并重新预览；
  `excel_autofix`/`excel_task` 结果渲染为修复前→修复后差异表。
- 右侧 Web 面板升级为真 Excel 网格：客户端模块内嵌 `x-data-spreadsheet`
  （MIT，列标/行号/公式栏风格），`excel_preview` 返回结构化 `sheets` 供
  网格渲染；实机验证网格渲染成功（截图 assets/panel-preview.png）。
- 就地实时编辑（M5）：插件注册 `/excel-set`、`/excel-undo` 命令，客户端经
  `remote.commands` 直调——改单元格回车即**就地写本地文件**（自动备份
  `.bak` + patch 审计 + 公式体检），面板提供“撤销本次修改”按钮按审计回滚；
  `live-edit.test.ts` 覆盖磁盘级写回与恢复。
- 修复 #1：DSH 宿主包（`dsh-tools` / `dsh-llm` / `dsh-system-prompt` /
  `dsh-attachment`）从 `dependencies` 移至 `peerDependencies`
  （`^0.1.0-rc.6`，与宿主版本对齐），根包改为 devDependencies——避免装进
  profile 后遮蔽宿主 rc.6 造成 Symbol 身份冲突（工具调用失败 / 极简模式
  挂载失败）；新增打包回归测试：bundle 的 dependencies 不得含任何
  `@deepseek-ai/dsh-*` 宿主包。

## v0.33.0 — 2026-08-15

- 新增 `excel_task` 多步任务编排：一次调用执行一串 `excel_operate` 步骤，
  每步结束自动体检公式、确定性修复后进入下一步，输出最终文件。
- 新增 `excel_explain_formula` 公式解释：把函数、引用区域、跨表引用、
  运算翻译成人话。
- `excel_operate` 新增 `normalizeText` 文本标准化（全角→半角、去重空格）。
- 测试规模 165 → 171。

## v0.32.0 — 2026-08-15

- `excel_operate` 新增 `highlightRows` 整行条件高亮（按多条件匹配整行并
  填充样式，默认黄色）。
- 新增 `excel_insight` 数据洞察工具：纯启发式数据体检（缺失/重复/异常值/
  负值/空格/公式），不依赖 LLM，并给出下一步建议。
- `excel_operate` 新增 `fuzzyMatch` 两表模糊匹配（按键相似度匹配并回填
  目标值，可输出匹配分数）。
- 测试规模 161 → 165。

## v0.31.0 — 2026-08-14

- 解决“用户不知道怎么描述”的入口问题：
  - 新增 `excel_menu`：给文件就出菜单——一句话摘要（行数/列数/表头/空值/
    公式）+ 清洗、补空值、报表、透视、图表、体检、通知、岗位模板等可选
    方案，每个带示例话术，把“描述需求”变成“选需求”。
  - 注册系统提示交互策略段：业务目标优先、模糊就给选项、先做后改并提示
    `excel_undo` 回滚、破坏性操作先确认、岗位模板一键套用。
- 测试规模 159 → 161。

## v0.30.0 — 2026-08-14

- `excel_operate` 新增 7 个数据清洗操作：`dedupeRows`（按列去重、保留首/末行）、`fillMissing`（固定值/向上/
  向左填充）、`removeEmptyRows` / `removeEmptyColumns`、`trimText`、
  `changeCase`（upper/lower/proper）、`splitColumn`（按分隔符分列，自动插入
  右侧新列并联动公式引用）。
- 新增生态能力对照文档：逐仓库记录核心效果、整合状态与缺口，作为后续
  迭代路线图（文档已在 v0.34.1 后移除，定位转向原创能力与评测驱动）。
- 测试规模 152 → 159。

## v0.29.0 — 2026-08-14

- 新增 `excel_profile` 大表速览：结构化表格编码（表头识别、每列类型/缺失/
  唯一值/数值区间/日期范围/高频值/样例 + 建议读取范围）；`excel_read` 新增
  `maxRows` 分页读取，避免整表灌入对话爆 token。
- 新增 `excel_autofix` 一键自愈闭环：体检 → 确定性修复（可选 LLM）→ 复检 →
  人话汇报，一条命令完成“检查并修复”。
- 测试规模 147 → 152。

## v0.28.0 — 2026-08-14

- 新增 `importCsv` / `exportCsv`（RFC 4180，支持分隔符与引号转义）；导出
  默认开启公式注入防护（值以 `= + - @` 开头时加 `'` 前缀）。
- 修复：异步操作（importCsv/exportCsv）此前未被 await，工作簿先写盘导致竞态；
  改为顺序等待。
- 测试规模 146 → 147。

## v0.27.0 — 2026-08-14

- 条件格式补齐职场常用规则：重复值/唯一值、空值/非空、错误/无错误、高于/低于
  平均、日期周期（今天/昨天/本周/本月等）、不包含文本。ExcelJS 不支持的原生
  类型自动翻译为等价的 expression 公式（COUNTIF / ISBLANK / ISERROR / SEARCH），
  保证写入后真实生效。
- 测试规模 145 → 146。

## v0.26.0 — 2026-08-14

- 新增 `preset` 岗位模板操作：运营（ops）= 一键报表 + 数据条；产品（product）=
  一键报表 + 色阶；数分（data）= 一键报表 + 色阶 + 按条件筛选明细副本。
  一条指令按岗位出整套活。
- 新增 [docs/roles.md](docs/roles.md) 岗位指南：三个岗位的典型任务、preset 参数、
  对话示例与进阶组合。
- 测试规模 144 → 145（场景6：三岗位 preset 端到端）。

## v0.25.0 — 2026-08-14

- 新增 `report` 组合模板操作：一条 op 完成“排序 → 分类汇总 → 动态 SUMIFS 透视
  汇总表 → 自动筛选 → 表头样式 → 冻结首行 → 数字格式”，把多轮对话分析压成一次
  调用；顺序设计保证汇总公式覆盖最终数据块。
- 测试规模 143 → 144（场景5：一键产出经营报表端到端）。

## v0.24.1 — 2026-08-14

- `subtotal` 插入小计/总计行后，同表与跨表的**相对**引用会像 Excel 一样联动平移
  （此前只移动行、不改引用）。绝对引用（如 `aggregateReport` 生成的
  `$B$2:$B$7`）按 Excel 语义不自动扩展，属预期行为。
- 真实对话回归（5 场景）通过：报表搭建 4 轮、修复 4 轮、分析 6 轮、VLOOKUP+合并
  4 轮、图表 2 轮。
- 测试规模 142 → 143。

## v0.24.0 — 2026-08-14

- 兼容修复：含原生数据透视表的文件此前会让 ExcelJS 读取崩溃，现在读写/校验/
  操作前自动剥离透视锚点（`stripPivotTableParts`），透视文件可正常读、查、改
  （重写时透视部件不保留，属 ExcelJS 限制）。
- 误报修复：`SUBTOTAL` 分类汇总行不再被当作列 pattern 异常，操作后验证更干净。
- 公式函数扩展：新增 `IF`、`XLOOKUP`、`CONCATENATE`、`LEFT`、`RIGHT`、`MID`、
  `ROUNDUP`、`ROUNDDOWN`，与 LLM 修复顾问同步支持。
- 测试规模 139 → 142。

## v0.23.1 — 2026-08-14

- 正式对外包装：README 与 npm 页新增“给使用者：一分钟上手”（安装命令、
  对话示例、平台说明、版本锁定）与相关链接；dsh-eval 仓库转公开并上主题标签。

## v0.23.0 — 2026-08-14

- 原生透视扩展：`excel_create_pivot` 支持多个行字段、列字段、报表筛选器，
  叠加原有的多值字段（求和/计数/平均/最大/最小），达到日常透视用法。
- 测试规模 138 → 139（多行 + 列字段端到端：Excel COM 生成 → 拆包验证
  `rowFields count="2"` / `colFields count="1"`）。

## v0.22.0 — 2026-08-14

- 原生数据透视表：新增 `excel_create_pivot` 工具，驱动 Excel COM 生成真正的
  pivotCache + pivotTable（不是模拟汇总），支持一个行字段 + 多个值字段
  （求和/计数/平均/最大/最小），Excel 内可刷新、可继续交互。Windows + Excel。
- 测试规模 137 → 138（pivot COM 端到端：生成 → 拆包验证 → Excel 打开）。

## v0.21.1 — 2026-08-14

- 修复 `excel_read` 工具级调用报错（dsh 要求无损 JSON，可选 undefined 字段被
  序列化丢弃），并新增插件级回归测试。
- 真实对话端到端扩到 5 场景：报表搭建、公式修复、数据分析（透视汇总 + 分类
  汇总 + 数据条 + 保护）、VLOOKUP + 邮件合并、图表创建。实测全部通过：
  4/4/6/4/2 轮对话完成。
- 测试规模 136 → 137。

## v0.21.0 — 2026-08-14

- 操作审计与撤销：`excel_operate` 每次编辑自动写 `<out>.patch.json`（单元格级
  前后差异），新增 `excel_undo` 工具按日志回滚（带前置条件校验，内容级撤销；
  行列结构变化不还原但单元格内容恢复）。
- 真实工作场景回归测试（`tests/workplace-scenarios.test.ts`）：用接近实际的
  销售台账（订单 + 产品价目 + 模板）串联四套场景——按区域动态透视报表、
  VLOOKUP 补产品名称、排序 + 分类汇总、邮件合并 + 工作表保护 + 命名区域。
- 测试规模 131 → 136。

## v0.20.0 — 2026-08-14

精细化操作层：

- 新增 `excel_read` 工具：精确读取单元格的值、公式、类型、数字格式、字体/填充/
  对齐、合并范围与数据有效性——编辑前模型能看清每一个单元格的准确状态。
- `style` 扩展：字号、字体名、四边边框（线型 + 颜色）。
- `conditionalFormatting` 扩展：数据条、色阶、图标集、包含文本、前 N 项。
- `protectSheet` 细化：可选择允许的单元格选择/格式/插入删除行列/排序/筛选等权限。
- 新增 `pageSetup`（打印区域、方向、缩放、页边距、居中）与 `definedName`（命名区域）。
- 测试规模 125 → 131。

## v0.19.0 — 2026-08-14

覆盖职场 Excel 15 讲（数据透视表 → 邮件合并）：

- Formula IR 泛化：新增 `function` 操作与 `range` 操作数，支持 VLOOKUP、INDEX、
  MATCH、ROUND、TEXT、SUMIF、COUNTIF、AVERAGE、MEDIAN、MAX、MIN、COUNT、COUNTA、
  日期函数（TODAY/YEAR/MONTH/DAY/DATE/DATEDIF/EOMONTH）与 SUMIFS/AVERAGEIFS/
  COUNTIFS；LLM 修复顾问与 `excel_compile_formula` 同步支持。
- `excel_operate` 新增：
  - `subtotal` — 分类汇总（SUBTOTAL 公式 + 分组小计 + 总计，粗体样式）。
  - `aggregateReport` — 动态数据分析报表：按分组列生成透视式汇总表，指标用
    实时 SUMIFS/AVERAGEIFS/COUNTIFS/MAXIFS/MINIFS 公式，源数据变化自动更新。
  - `filterToRange` — 高级筛选：按多条件把匹配行写到指定区域。
  - `protectSheet` / `unprotectSheet` — 工作表保护（可设密码）。
  - `mailMerge` — 邮件合并（Excel 侧）：模板 `{占位符}` 按数据行批量展开。
- 图表可视化：`excel_create_chart` / `excel_modify_chart`（Excel COM，Windows）——
  创建图表（类型/标题/数据范围）、修改参数（类型/标题/图例/坐标轴标题）。
- 测试规模 116 → 125。

## v0.18.0 — 2026-08-14

- 更名为 `dsh-excel-chat`：定位从“公式验证/修复工具”升级为“在 DeepSeek Harness
  里对话完成 Excel 工作”。npm 新包已发布，旧包 `dsh-excel-vera-plugin` 弃用并
  提示改名；GitHub 仓库同步改名（旧链接自动跳转）。
- README / bundle README / npm description 重写为对话优先：安装即聊，示例场景、
  能力清单、自动体检闭环。原名 VERA 保留为内部代号。

## v0.17.0 — 2026-08-14

- 对话直用打通：`excel_operate` 的 operations 改为严格的 27 操作判别联合 schema
  （`src/operation-schema.ts`），模型无需猜字段结构，首次调用即正确。
- `excel_repair_formulas` 新增 `outPath` 参数，修复结果可写到指定路径。
- DeepSeek 客户端支持原生函数调用（`deepseekChatWithTools`），新增
  `tests/invoke-conversation.ts` 真实对话端到端：自然语言 → 模型调工具 →
  执行 → 复验。实测两场景通过：报表搭建（合计列+加粗+冻结+筛选）2 轮、
  公式静默错误检测修复 3 轮。
- 测试规模 115 → 116。

## v0.16.0 — 2026-08-14

- `excel_operate` 新增报表骨架能力：
  - `autoFilter`：一键给表头区域加筛选下拉。
  - `addTable`：把区域转成结构化表格（Ctrl+T 效果），自动读取表头与行数据，
    支持斑马纹/表头行/汇总行。
- 错误值体检：`excel_validate_formulas` 新增 `error-value` 检测，扫描
  `#REF!` / `#DIV/0!` / `#VALUE!` / `#NAME?` / `#N/A` / `#NULL!` / `#NUM!`。
- 测试规模 112 → 115。

## v0.15.0 — 2026-08-14

- `excel_operate` 继续扩展：
  - `sortRange`：按一个或多个键排序（升序/降序），支持跳过表头行，稳定排序。
  - `dataValidation`：下拉列表（`list`）与数值/日期/文本长度校验（`between` 等
    运算符、错误提示、允许空值）。
  - `conditionalFormatting`：`cellIs` / `expression` 条件格式（如“大于 80 标红”）。
  - 删除行列时，引用被删单元格的公式现在转成 `#REF!`，与 Excel 原生行为一致；
    修复了删除后引用误判为 `#REF!` 的顺序问题。
- 测试规模 108 → 112。

## v0.14.0 — 2026-08-14

- `excel_operate` 扩展到职场级：
  - 值类型识别：`set`/`fill` 自动把数字、日期、布尔写成真正的类型，不再把
    `100` 写成文本（文本数字会让 SUM 类公式静默失效）。
  - 列操作：`insertColumns` / `deleteColumns`，公式列引用联动（含跨表引用），
    删除列时警告引用了被删列的公式。
  - 数据操作：`copyRange` / `moveRange`（公式按目标偏移调整）、`fillSeries`
    （数字/日期序列）。
  - 格式与视图：`style`（粗体/斜体/下划线/字体色/填充/数字格式/对齐/自动换行）、
    `setColumnWidth` / `setRowHeight` / `freezePanes`。
  - 查找与工作表：`findReplace`（大小写可选，返回替换次数）、`duplicateSheet`、
    `hideSheet`、`setTabColor`。
- 测试规模 98 → 108，覆盖全部新操作与值类型。

## v0.13.0 — 2026-08-14

- 新增 `excel_operate` 操作工具（面向日常 Excel 用户）：
  - `set` 写值/公式、`fill` 拖拽填充（自动平移相对行列引用）、
    `insertRows` / `deleteRows` 插入删除行（公式引用像 Excel 一样联动，含
    跨表引用）、`addSheet` / `renameSheet`（引用同步更新）/ `deleteSheet`、
    `clear` 清空、`merge` / `unmerge` 合并单元格。
  - 每次操作后自动复验公式静默错误，返回 `validation` 结果；删除行时提示
    引用了被删行的公式。
- 测试规模 85 → 98，新增 operations 单元测试与插件级 `excel_operate` 调用测试。

## v0.12.0 — 2026-08-14

- 包装对齐社区 dsh-plugin 最佳实践：npm 元数据补全（`exports` / `types` /
  `engines` / `keywords` / `peerDependencies` / `publishConfig` /
  `prepublishOnly`），bundle 构建生成 `dist/index.d.ts` 类型声明。
- 自动发布流水线：`.github/workflows/publish.yml` 在 `v*` tag 推送时执行
  测试 → 构建 → 打包（校验 tag 与版本一致）→ npm 发布（`NPM_TOKEN`
  自动化 token，绕过 2FA）→ 创建 GitHub Release 并附带 tarball。
- GitHub 仓库添加 `dsh-plugin`、`deepseek-harness` 等主题标签，进入
  [dsh-plugin topic](https://github.com/topics/dsh-plugin) 生态。
- README 补全：安装方式、7 个工具表、npm/Release 徽章、CI 发布说明。
- 打包测试新增断言：tarball 必须包含类型声明与 `cordis.patch.yml` 导出。

## v0.11.0 — 2026-08-14

- 空行确定性修复：`empty-gap` 自动克隆相邻公式填充缺口，`shiftFormulaRow`
  只平移相对行引用，保留绝对行（`$4`）与跨表前缀；补丁抽象支持空单元格
  （`oldValue: ''` 即“缺失”）。
- Oracle 闭环：`excel_repair_formulas` 新增 `oraclePath` 参数，修复后自动对比
  ground-truth workbook，结果中返回 `oracleScore`（准确率 + mismatch 明细）。
- Benchmark 扩展到 11 个任务（`src/benchmark-cases.ts`）：范围双端点、绝对引用、
  空行填充、跨表、多表、聚合结构、hardcode、结构不匹配等。
- LLM 容错：aggregate SUM 缺 metric 时从表结构第一列推断；单条 malformed IR
  跳过而不中断整个修复；benchmark 记录 `llmError` 而非崩溃。
- Prompt 增强：加入“列 pattern 示例公式 + aggregate 示例”，要求模型对齐
  示例公式形态。真实 DeepSeek：Pass@1 11/11，meanAccuracy 1.000。
- 测试规模 75 → 82。

## v0.10.0 — 2026-08-14

- 范围尾引用修复：确定性 repair 现在覆盖 `=SUM(B4:C3)` 这类 range.end 偏移异常；
  两个端点同时偏移时一次重建整段范围，未偏移端点保留原文本（含 `$` 绝对修饰）。
- 自动表头检测：`detectTableFromCells` 从单元格内容推断 `{ sheet, columns }`；
  `excel_repair_formulas` 新增 `autoTable` 参数，`useLlm` 时无需手写 table schema。
- Oracle 判分：`scoreWorkbookAgainstOracle` 按单元格对比候选与标准 workbook，
  容忍公式大小写/空白与数字格式差异，输出准确率与 mismatch 明细。
- Pass@1 Benchmark：`runBenchmark` 按“确定性修复 → LLM 修复”真实流程执行任务，
  与 oracle 对比给出 Pass@1 与平均准确率；`tests/invoke-benchmark.ts` 可接真实 DeepSeek。
- 测试规模 56+ 增至 75+，覆盖自动表头、范围修复、判分与 benchmark。

## v0.9.0 — 2026-08-14

- 包名确定为 `dsh-excel-vera-plugin`（dsh 生态前缀 + Excel 品牌），
  `cordis.patch.yml` 与 bundle README 同步更新，发布前改名。

## v0.8.0 — 2026-08-14

- 真实 DeepSeek 模型端到端：`deepseekChatCompletion` / `deepseekLlmTextFromEnv` 直接调用
  chat completions（读 `DEEPSEEK_API_KEY`），接到修复顾问后完成 读取 → 验证 → LLM 生成 IR →
  编译 → Patch → 复验 全流程；`invoke-real-llm` 已用真实 API 跑通
  （`Sheet1!D3: =SUM(B3:C3) → =B3-C3`，复验异常归零）。
- 结构异常检测改进：只标记缺少“多数派槽位”的单元格，D2/D4 正常、D3 异常时不再误报多数派。
- LLM 输出容错：operand 裸字符串自动归一化为 cell/column；裸单元格 id 自动匹配
  sheet 限定键（`D3` → `Sheet1!D3`）；prompt 内置 IR 示例和“只修真正偏离的单元格”约束。

## v0.7.0 — 2026-08-14

- 视觉评审接入真实模型链路：`visionTextFromContext` 用 `ctx.attachments` 上传图表 PNG、
  构造带图片块的用户消息、走 `ctx.llm.stream`；新工具 `excel_validate_charts_visual`。
- 测试用假 attachment store + 假 adapter 验证完整链路（图片块确实随消息发出）。

## v0.6.0 — 2026-08-14

- Chart Visual Validator 骨架：`exportChartsWithExcel` 用本地 Excel COM 把图表导出为 PNG；
  `createVisionCritic` 把视觉 checklist（标题截断 / 图例遮挡 / 标签重叠 / 坐标轴 / 拥挤 / 趋势可读性）
  交给可注入的视觉函数，假实现可完整测试。新工具 `excel_export_charts`。
- bundle 发布准备：MIT LICENSE、bundle README、移除 private；`npm pack --dry-run` 验证
  tarball 包含 `dist/`、`cordis.patch.yml`、README、LICENSE。

## v0.5.0 — 2026-08-14

- Chart Semantic Validator：直接解析 xlsx 内的 chart XML（fflate），结构校验包括
  图表类型、系列引用存在性、缺失单元格、二维范围、日期未排序；新工具 `excel_validate_charts`。
- 可发布 bundle：`bundle/` 目录包含 `dsh.bundle` 清单、`cordis.patch.yml` 和 tsc 编译产物，
  `dsh plugin --profile <name> add ./bundle` 可直接安装；git 安装时 `prepare` 自动构建。
- 全部工具汇总：`excel_validate_formulas` / `excel_compile_formula` / `excel_repair_formulas` /
  `excel_diff_workbook` / `excel_validate_charts`。

## v0.4.0 — 2026-08-14

- `ctx.llm` 接线：`llmTextFromContext` 把 dsh 流式 LLM 服务包装成修复顾问的 `LlmText`，
  `excel_repair_formulas` 新增 `useLlm` / `provider` / `model` / `table` 参数，
  可选注入（`ctx.get('llm')`），无 LLM 服务时确定性修复照常工作。
- Workbook Diff：`diffCellMaps` / `diffWorkbookFiles` 输出 added / removed / changed 单元格，
  新工具 `excel_diff_workbook`（给 Excel 的 Git Diff）。
- Patch Log：`writePatchLog` / `readPatchLog` / `applyPatchLog` / `rollbackPatchLog`，
  修复可审计、可回滚。
- LLM 接线测试使用假 adapter 走完整 `ctx.llm.stream` 管线，无需 API key。

## v0.3.0 — 2026-08-14

- Formula IR JSON Schema：`excel_compile_formula` 的 `ir` 参数改用严格 `oneOf` schema，
  模型输出非法 IR 时在参数校验阶段即被拒绝。
- LLM 修复顾问：`createLlmRepairAdvisor` 把 workbook 摘录 + 异常列表 + 表结构组装成 prompt，
  让 LLM 返回 IR 修复，再走 Compiler → Patch → 复验；LLM 以函数注入，无 API key 也可完整测试。

## v0.2.0 — 2026-08-14

- Formula Patch 抽象：apply / revert / 写回 workbook，带前置条件校验。
- 确定性 Repair：reference-offset 异常 → 生成最小补丁（如 `Sales!D4: =B4-C3 → =B4-C4`）。
- 新工具 `excel_repair_formulas`：读取文件 → 验证 → 修复 → 写出 `.repaired.xlsx` → 复验。
- 公式解析器记录引用文本范围（`range`），支持安全替换。

## v0.1.0 — 2026-08-14

- P0 Formula Pattern Validator：A1 引用解析、依赖图、列 pattern 异常检测、hardcode / 空行 / 循环引用检测。
- Formula IR + Compiler：binary / ratio / aggregate 语义 IR → 确定性 Excel 公式。
- ExcelJS workbook 读取，真实 .xlsx 文件端到端验证。
- dsh 插件入口：`excel_validate_formulas`、`excel_compile_formula`。
