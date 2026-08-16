# Changelog

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
