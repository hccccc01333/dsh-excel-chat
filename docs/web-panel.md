# 右侧 Excel 面板（Web UI）方案

## 结论：可行，且安装插件即可获得

DeepSeek Harness Web 支持插件自带**前端客户端模块**：

- 插件包在 package.json 声明 `dsh.client`，并提供 `exports["./client"]`；
- web shell 启动时加载 `/plugins/<id>/client.js` 并运行其 `apply(ctx)`；
- 客户端插件可注入 UI slot、订阅会话事件、经 `connection` 调用 host 侧能力。

web 布局插件（`dsh-client-ui-layout`）声明了**右侧详情列 slot `details`**
（`scope: session`），当前由 ui-conversation 的 DetailsPanel 占据，并声明了
内层座位（如 tool-details）。这就是“右侧展示/修改 Excel”的挂载点。

## 目标体验

```
┌────────────┬───────────────────┬──────────────────┐
│ 会话列表    │ 对话（左侧）        │ Excel 面板（右侧） │
│            │  用户/agent 消息    │  工作表标签页      │
│            │  工具行（excel_*）  │  表格网格（可编辑） │
│            │                   │  公式体检结果      │
└────────────┴───────────────────┴──────────────────┘
```

用户安装 dsh-excel-chat 后，右侧自动出现“工作表”面板：

1. 会话里出现 `excel_read` / `excel_preview` / `excel_operate` 调用时，面板
   自动定位到当前工作簿并渲染；
2. 只读模式：显示工作表标签 + 表格（值/公式/类型），数据来自
   `excel_read` / `excel_preview`；
3. 编辑模式：单元格编辑 → 组装 `excel_operate.set` → 经 host RPC 执行 →
   面板刷新 + 展示公式体检结果（复用 excel_validate_formulas）；
4. 与 `excel_task` / `excel_autofix` 结果联动：面板展示修复前后差异。

## 架构

- 客户端模块：`src/client/`（React + slots + connection），随 npm 包分发
  （`exports["./client"]` + `dsh.client`）；
- 数据通道：面板 → `connection.call(channel, endpoint, payload)` → host 执行
  excel 工具 → 返回表格数据/编辑结果；
- 事件源：订阅会话的工具调用事件，或由工具输出携带结构化表数据（新增
  `excel_read` 的可选 `render` 负载）。

## 待确认的 API（实现前）

- `details` 列内层 seat 的准确名称与注册方式（additive 而非替换整列）；
- host 侧工具执行的 RPC channel/endpoint（`dsh-client-connection` 的
  `call(channel, endpoint, payload)` 契约）；
- 客户端构建工具链（tsdown client 配置、React 依赖如何随 npm 包分发）。

## MVP 里程碑

1. 只读预览：右侧面板显示当前工作簿的表格（工具调用后自动出现）
2. 单元格编辑：编辑 → `excel_operate.set` → 刷新 + 体检
3. 与 goal/task/autofix 结果联动（修复差异视图）
4. 随 npm 包分发 + harness web 集成测试（复用 .demo-recorder 回放）

## 边界

- 面板属 harness Web 前端能力；插件负责提供 client 模块与数据契约。
- 服务端工具与验证逻辑不变（excel_operate / validate / autofix 已在插件内）。
