# Source modules

What each file does. This is an inventory for people changing the project, not usage
documentation — see [README](../README.en.md) for that, including the tool list.

- `src/formula.ts` — A1 reference parser (cell, range, cross-sheet, whole-column), canonical cell ids, column helpers.
- `src/graph.ts` — dependency graph with bounded range expansion and cycle detection.
- `src/patterns.ts` — per-column reference-pattern analysis: offset anomalies, structure mismatches, hardcode breaks, empty gaps.
- `src/validator.ts` — `validate(cells)` entry point returning the graph, column reports and anomalies.
- `src/ir.ts` — Formula IR types (binary / ratio / aggregate).
- `src/ir-schema.ts` — the dsh tool-DSL schema for Formula IR (strict oneOf validation).
- `src/compiler.ts` — `compileFormula(ir, { baseCell, table })` compiles to an Excel formula.
- `src/advisor.ts` — LLM repair advisor: anomalies + table structure → prompt → IR repair → Patch.
- `src/llm.ts` — `llmTextFromContext`: wires the `ctx.llm` streaming service into the repair advisor (optional injection).
- `src/diff.ts` — workbook diff and patch log: diff / apply / rollback.
- `src/charts.ts` / `src/chart-validator.ts` — xlsx chart XML parsing and structure validation.
- `src/chart-visual.ts` — Excel COM chart creation / parameter edits / export, plus an injectable visual review (VLM interface).
- `src/vision.ts` — `visionTextFromContext`: wires `ctx.attachments` + `ctx.llm` into a visual review.
- `src/deepseek.ts` — DeepSeek chat completions client (reads `DEEPSEEK_API_KEY`), used by the repair advisor.
- `src/patch.ts` — minimal patch abstraction: apply / revert / write back to a workbook.
- `src/repair.ts` — turns validation results into deterministic repairs (reference offsets and empty-row fills), writes `.repaired.xlsx` and re-validates; optionally takes oracle cells and returns `oracleScore`.
- `src/workbook.ts` — ExcelJS-based workbook reader: `.xlsx` → cell-content map, plus `validateWorkbookFile(path)`.
- `src/tables.ts` — `detectTableFromCells`: infers `{ sheet, columns }` from cell contents, backing `excel_repair_formulas`' `autoTable` header detection.
- `src/score.ts` — `scoreWorkbookAgainstOracle`: oracle scoring at cell level, tolerating formula case/whitespace and number-format differences, returning accuracy and mismatch details.
- `src/read.ts` — `readWorkbookDetail`: exact cell reads (value / formula / type / format / merges / data validation), used by the `excel_read` tool.
- `src/profile.ts` — `profileWorkbook`: structured table encoding, producing a compact per-sheet, per-column profile and a suggested read range for the `excel_profile` tool.
- `src/autofix.ts` — `autofixWorkbookFile`: the check → repair → re-check → plain-language summary self-healing loop behind `excel_autofix`.
- `src/pivot.ts` — `createPivotTable`: drives Excel COM to generate a native pivot table (pivotCache + pivotTable) that always opens cleanly.
- `src/operation-schema.ts` — the strict discriminated-union schema for `excel_operate`'s 77 operations, so the model emits the right shape straight from the `op` field.
- `src/operations.ts` — the Excel operation DSL: set (type inference) / fill / fillSeries / insertRows / deleteRows / insertColumns / deleteColumns (references shift like Excel, including across sheets, and deleted cells become `#REF!`) / sortRange (multi-key) / copyRange / moveRange / style / dataValidation (dropdowns and numeric checks) / conditionalFormatting / setColumnWidth / autoFilter / addTable (structured tables) / setRowHeight / freezePanes / findReplace / addSheet / renameSheet / deleteSheet / duplicateSheet / hideSheet / setTabColor / clear / merge / unmerge.
- `src/benchmark.ts` — Pass@1 benchmark: deterministic repair → LLM repair, scored against an oracle.
- `src/benchmark-cases.ts` — 11 benchmark tasks: range endpoints, absolute references, empty rows, cross-sheet, multi-sheet, aggregate structure, hardcode and more.
- `src/file-benchmark.ts` + `src/corpus/` — ExcelBench lite: 113 file-level real workplace tasks (editing / analysis / formulas / workflows); see [docs/benchmark.md](docs/benchmark.md).
- `src/index.ts` — the dsh plugin entry exposing 25 tools (understanding a file / formula checks and repair / operations and orchestration / charts and export; the full list is in the Tools table above).
- `bundle/` — the publishable dsh bundle: manifest + cordis.patch.yml + compiled output.
