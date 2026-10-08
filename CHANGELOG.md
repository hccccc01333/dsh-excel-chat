# Changelog

## Unreleased

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

- **新增 5 个测试文件 / 46 例**，补上此前没有直接测试的模块：
  `tests/csv.test.ts`（12 例，含端到端 `importCsv` 回归）、
  `tests/graph.test.ts`（14 例）、`tests/charts.test.ts`（8 例）、
  `tests/formula.test.ts`（10 例，列字母与单元格 id 的边界校验）、
  `tests/operations.test.ts` 的 `a column name where a column letter belongs is
  rejected, not allocated`（端到端 OOM 回归），以及
  `tests/file-benchmark.test.ts` 的语料守卫
  `no corpus fixture ships a circular formula`——遍历 100 个夹具断言无环，
  这条守卫当初就能抓到那两处夹具缺陷。测试 350 → **396 通过**。

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
