window.__ModuleLoader__.load({
	id: "dsh-excel-chat",
	factory: (require) => {
		var module = { exports: {} };
		var exports = module.exports;
		Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });
		let react = require("react");
		//#region src/client/index.tsx
		/**
		* dsh-excel-chat browser half: registers per-tool renderers into the keyed
		* `tool.call.toolview` seat, so selecting an excel_* tool row in chat shows a
		* rendered table / summary in the right details column instead of raw JSON.
		*
		* M2 editing: cells are editable; committing an edit fills the composer draft
		* through `inputActions.setDraft` and submits it, so the agent applies the
		* change via excel_operate.set and re-previews. M3: excel_autofix / excel_task
		* results render as an old→new repair diff.
		*/
		/** Tools whose settled result is a cell table (excel_read / excel_preview). */
		const TABLE_TOOLS = new Set(["excel_read", "excel_preview"]);
		/** Tools whose settled result is a report/summary object. */
		const SUMMARY_TOOLS = new Set([
			"excel_insight",
			"excel_operate",
			"excel_autofix",
			"excel_task",
			"excel_menu"
		]);
		/** Tools whose result carries repair diffs (old → new). */
		const DIFF_TOOLS = new Set(["excel_autofix", "excel_task"]);
		const inject = ["slots"];
		function cellText(value, formula) {
			if (formula) return formula;
			if (value === null || value === void 0) return "";
			if (typeof value === "object") return JSON.stringify(value);
			return String(value);
		}
		function tableCellStyle(header) {
			return {
				border: "1px solid #ddd",
				padding: "4px 8px",
				textAlign: "left",
				background: header ? "#f3f4f6" : void 0,
				fontSize: 12
			};
		}
		/** Minimal table renderer for excel_read/excel_preview results (M1 + M2 editing). */
		function TableFromResult(result, editable, editing, onEditStart, onEditCommit) {
			const value = result;
			const sheets = Array.isArray(value.sheets) ? value.sheets : null;
			if (!sheets) {
				if (typeof value.markdown === "string") return (0, react.createElement)("pre", { style: {
					fontSize: 12,
					overflow: "auto",
					whiteSpace: "pre-wrap",
					maxHeight: 480
				} }, value.markdown);
				return null;
			}
			const first = sheets[0];
			if (!first) return null;
			const columns = /* @__PURE__ */ new Set();
			for (const cell of first.cells) {
				const match = /([A-Za-z]+)\d+$/.exec(cell.id.split("!").pop() ?? "");
				if (match) columns.add(match[1]);
			}
			const columnList = [...columns].sort((a, b) => a.length - b.length || a.localeCompare(b));
			const byCell = new Map(first.cells.map((cell) => [cell.id, cell]));
			const rows = /* @__PURE__ */ new Set();
			for (const cell of first.cells) {
				const match = /(\d+)$/.exec(cell.id.split("!").pop() ?? "");
				if (match) rows.add(Number(match[1]));
			}
			const rowList = [...rows].sort((a, b) => a - b);
			const renderCell = (column, row) => {
				const id = `${first.sheet}!${column}${row}`;
				const cell = byCell.get(id);
				const text = cell ? cellText(cell.value, cell.formula) : "";
				if (editable && editing?.id === id) return (0, react.createElement)("input", {
					defaultValue: editing.value,
					autoFocus: true,
					onKeyDown: (event) => {
						if (event.key === "Enter") onEditCommit({
							id,
							value: event.target.value
						});
						if (event.key === "Escape") onEditStart({
							id: "",
							value: ""
						});
					},
					onBlur: (event) => onEditCommit({
						id,
						value: event.target.value
					}),
					style: {
						width: "100%",
						boxSizing: "border-box"
					}
				});
				return (0, react.createElement)(editable ? "div" : "span", editable ? {
					onDoubleClick: () => onEditStart({
						id,
						value: text
					}),
					title: "双击编辑，回车提交（通过对话让 agent 执行 set）",
					style: {
						cursor: "pointer",
						minHeight: 18
					}
				} : null, text);
			};
			return (0, react.createElement)("div", { style: {
				overflow: "auto",
				maxHeight: 480
			} }, (0, react.createElement)("div", { style: {
				fontWeight: 600,
				marginBottom: 6
			} }, first.sheet), (0, react.createElement)("table", { style: {
				borderCollapse: "collapse",
				width: "100%"
			} }, (0, react.createElement)("thead", null, (0, react.createElement)("tr", null, columnList.map((column) => (0, react.createElement)("th", {
				key: column,
				style: tableCellStyle(true)
			}, column)))), (0, react.createElement)("tbody", null, rowList.map((row) => (0, react.createElement)("tr", { key: row }, columnList.map((column) => (0, react.createElement)("td", {
				key: `${column}${row}`,
				style: tableCellStyle(false)
			}, renderCell(column, row))))))));
		}
		function settledText(block) {
			if (!("kind" in block)) return null;
			return block.content.map((item) => item.type === "text" ? item.text : JSON.stringify(item)).join("\n") || null;
		}
		/** M3: render repair diffs (old → new) from excel_autofix / excel_task results. */
		function DiffFromResult(result) {
			const record = result;
			const repairs = Array.isArray(record.repairs) ? record.repairs : [];
			const rows = [];
			for (const repair of repairs) {
				if (typeof repair.id !== "string") continue;
				rows.push({
					id: repair.id,
					old: String(repair.oldValue ?? ""),
					next: String(repair.newValue ?? "")
				});
			}
			const stepRows = [];
			if (Array.isArray(record.steps)) for (const step of record.steps) {
				const validation = step.validation;
				if (validation) stepRows.push(`${String(step.name ?? "")}: 修复前 ${String(validation.before ?? "?")} 异常 → 修复后 ${String(validation.after ?? "?")}（修复 ${String(validation.fixed ?? 0)} 处）`);
			}
			const lines = [];
			if (typeof record.message === "string") lines.push(record.message);
			if (typeof record.repairedPath === "string") lines.push(`输出：${record.repairedPath}`);
			if (typeof record.outputPath === "string") lines.push(`输出：${record.outputPath}`);
			const nodes = [];
			if (lines.length > 0) nodes.push((0, react.createElement)("pre", {
				key: "msg",
				style: {
					fontSize: 12,
					whiteSpace: "pre-wrap",
					margin: "0 0 8px"
				}
			}, lines.join("\n")));
			if (stepRows.length > 0) nodes.push((0, react.createElement)("pre", {
				key: "steps",
				style: {
					fontSize: 12,
					whiteSpace: "pre-wrap",
					margin: "0 0 8px"
				}
			}, stepRows.join("\n")));
			if (rows.length > 0) nodes.push((0, react.createElement)("table", {
				key: "diff",
				style: {
					borderCollapse: "collapse",
					width: "100%"
				}
			}, (0, react.createElement)("thead", null, (0, react.createElement)("tr", null, [
				"单元格",
				"修复前",
				"修复后"
			].map((label, index) => (0, react.createElement)("th", {
				key: label,
				style: tableCellStyle(true)
			}, label)))), (0, react.createElement)("tbody", null, rows.map((row) => (0, react.createElement)("tr", { key: row.id }, [
				(0, react.createElement)("td", {
					key: "id",
					style: tableCellStyle(false)
				}, row.id),
				(0, react.createElement)("td", {
					key: "old",
					style: {
						...tableCellStyle(false),
						textDecoration: "line-through",
						color: "#b91c1c"
					}
				}, row.old),
				(0, react.createElement)("td", {
					key: "new",
					style: {
						...tableCellStyle(false),
						color: "#15803d"
					}
				}, row.next)
			])))));
			return nodes.length > 0 ? (0, react.createElement)("div", null, nodes) : null;
		}
		/** Toolview for one excel_* tool: render table (editable), repair diff, or summary. */
		function ExcelToolView(props) {
			const { toolName, block, inputActions } = props;
			const [editing, setEditing] = (0, react.useState)(null);
			const [submitted, setSubmitted] = (0, react.useState)(null);
			const text = settledText(block);
			if (text === null) return null;
			let parsed = null;
			try {
				parsed = JSON.parse(text);
			} catch {}
			const canEdit = TABLE_TOOLS.has(toolName) && typeof inputActions?.setDraft === "function";
			if (TABLE_TOOLS.has(toolName) && parsed !== null) {
				const table = TableFromResult(parsed, canEdit, editing, (cell) => {
					setEditing(cell);
					setSubmitted(null);
				}, (cell) => {
					setEditing(null);
					if (!cell.id || !cell.value) return;
					const instruction = `请把 ${cell.id} 的值改成 ${cell.value}，用 excel_operate 的 set 执行，完成后用 excel_read 重新预览并回复“完成”。`;
					inputActions.setDraft(instruction);
					inputActions.submit();
					setSubmitted(`已提交修改：${cell.id} = ${cell.value}`);
				});
				if (table !== null) return (0, react.createElement)("div", null, table, submitted === null ? canEdit ? (0, react.createElement)("div", { style: {
					fontSize: 11,
					color: "#6b7280",
					marginTop: 6
				} }, "双击单元格编辑，回车提交给 agent 执行 set") : null : (0, react.createElement)("div", { style: {
					fontSize: 11,
					color: "#15803d",
					marginTop: 6
				} }, submitted));
			}
			if (DIFF_TOOLS.has(toolName) && parsed !== null && typeof parsed === "object") {
				const diff = DiffFromResult(parsed);
				if (diff !== null) return diff;
			}
			if (SUMMARY_TOOLS.has(toolName) && parsed !== null && typeof parsed === "object") {
				const record = parsed;
				const lines = [
					"message",
					"summary",
					"outputPath",
					"repairedPath"
				].filter((key) => typeof record[key] === "string").map((key) => `${key}: ${String(record[key])}`);
				const counts = ["before", "after"].filter((key) => record[key] && typeof record[key] === "object").map((key) => `${key}: ${JSON.stringify(record[key])}`);
				if (lines.length > 0 || counts.length > 0) return (0, react.createElement)("pre", { style: {
					fontSize: 12,
					overflow: "auto",
					whiteSpace: "pre-wrap",
					maxHeight: 480
				} }, [...lines, ...counts].join("\n"));
			}
			return (0, react.createElement)("pre", { style: {
				fontSize: 12,
				overflow: "auto",
				whiteSpace: "pre-wrap",
				maxHeight: 480
			} }, text);
		}
		/** Client plugin body: register the excel_* toolviews into the details column. */
		function apply(ctx) {
			for (const tool of [...TABLE_TOOLS, ...SUMMARY_TOOLS]) ctx.effect(() => ctx.slots.inject("tool.call.toolview", () => ctx.slots.register({
				name: "tool.call.toolview",
				key: tool
			}, ExcelToolView)), `dsh-excel-chat: toolview ${tool}`);
		}
		//#endregion
		exports.ExcelToolView = ExcelToolView;
		exports.apply = apply;
		exports.inject = inject;
		return module.exports;
	}
});

//# sourceMappingURL=client.js.map