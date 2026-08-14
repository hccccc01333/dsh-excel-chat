# Changelog

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
