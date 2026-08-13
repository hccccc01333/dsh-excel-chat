# dsh-vera-plugin

VERA（Verified Excel Reasoning Agent）的 dsh 插件 bundle。

安装：

```sh
dsh plugin --profile demo add ./bundle
```

工具：

- `excel_validate_formulas` — 公式静默错误检测（列 pattern、hardcode、空行、循环引用）
- `excel_compile_formula` — Formula IR → 确定性 Excel 公式
- `excel_repair_formulas` — 确定性修复 + 可选 LLM 修复，输出 `.repaired.xlsx` 并复验
- `excel_diff_workbook` — 两个 workbook 的单元格差异
- `excel_validate_charts` — 图表结构校验
- `excel_export_charts` — 用本地 Excel 把图表导出为 PNG（Windows）

构建（源码安装时 `prepare` 自动执行）：

```sh
npm run build
```
