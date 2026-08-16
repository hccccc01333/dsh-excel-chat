# 用户痛点 → 需求（Pain-Point-Driven Roadmap）

需求不按“想到什么功能加什么”，而从使用者真实痛点反推。本页每条痛点
都带证据（issue / benchmark / 同类仓库），不凭感觉。

## 痛点清单

| 痛点 | 证据 | 对应需求 | 优先级 |
|---|---|---|---|
| 不敢让 AI 动正式文件：说“完成”了结果却是错的 | v0.35 冒烟：分析类失败 3/5 是验证器误判（0/4、0/3、0/5 断言却判达成） | Verifier 2.0（断言级校验，不只查公式异常）；改动前后差异 + 一键回滚；健康报告写进文件 | P0 |
| 装完不能用，全部工具调用失败 | Issue #1：宿主包被当 dependencies 提升，Symbol 冲突 | 安装自检命令（`dsh-excel-chat doctor`）+ 装完自动跑一次冒烟调用；打包回归测试已加 | P0 |
| 公司数据不能出网 / API 太贵 | cellm #344：用户想接本地 Ollama | 支持 OpenAI-compatible / Ollama 本地模型 provider；路径沙箱与逐次确认已有 | P1 |
| 换台电脑（Excel 2021 LTSC / WPS / 中文版）就崩 | cellm #345：Excel 2021 LTSC zh-CN 弹 VBA 内存对话框；#218：不支持表格引用 | 兼容矩阵（Excel 2016/2021 LTSC / WPS）+ 结构化表格引用（`表[列]`）生成 | P1 |
| 不知道 AI 能干什么、描述不清 | 交互策略讨论：用户只会说“帮我整理一下” | `excel_menu` 菜单 + 示例话术已有；继续做成界面可点击示例，熟手可关闭引导 | P2 |
| 大表卡 / 爆 token | cellm #170：复制单元格卡顿；ExcelBench 部分任务要求分页 | `excel_profile` + `maxRows` 分页已有；补大表模式提示与流式确认 | P2 |
| 安全边界模糊：AI 能随便读写文件 | excel-mcp-server issue 群：命令注入、无路径限制、无鉴权 | 保留并强化 DSH 审批卡；文件路径策略可配置；只读模式 | P1 |

## 与既有路线图的关系

- v0.35 Failure Taxonomy：把“结果不可信”量化（已完成首块）
- v0.36 Semantic Layer：解决“找错列/表/指标”的语义错误
- v0.38 Verifier 2.0：直接打 P0 第一项
- v0.39 Excel Workspace：把“改动差异 + 回滚”变成可视闭环

不在痛点清单上的功能（Word/PPT 联动、MCP 适配等）继续推迟，直到有
真实用户证据说明需要。
