import { compileFormula } from './compiler.js';
import { columnToNumber, parseCellId } from './formula.js';
export function buildRepairPrompt(cells, anomalies, table, columns) {
    const example = {
        repairs: [{
                id: 'D3',
                baseCell: 'D3',
                ir: {
                    operation: 'binary',
                    left: { kind: 'column', column: 'revenue' },
                    right: { kind: 'column', column: 'cost' },
                    operator: '-',
                },
            }],
    };
    const aggregateExample = {
        repairs: [{
                id: 'D4',
                baseCell: 'D4',
                ir: {
                    operation: 'aggregate',
                    metric: 'revenue',
                    function: 'SUM',
                    filters: [],
                },
            }],
    };
    return [
        'You are the repair planner for a verified Excel agent.',
        'The workbook excerpt (cell id -> content) is:',
        JSON.stringify(cells, null, 2),
        'The validator found these anomalies:',
        JSON.stringify(anomalies, null, 2),
        ...(columns && columns.length > 0
            ? [
                'The column patterns are (first formula cell per column):',
                JSON.stringify(columns.map(columnExample(cells)), null, 2),
            ]
            : []),
        `The table schema is: ${JSON.stringify(table)}`,
        'Return ONLY JSON matching: {"repairs":[{"id":"<cell id>","baseCell":"<cell id>","ir":<FormulaIR>}]}.',
        'FormulaIR uses operation "binary" (left/right/operator), "ratio" (numerator/denominator),',
        'or "aggregate" (metric/function/filters, each filter with column and value_from).',
        'Operands are objects: {"kind":"column","column":"<logical column from table.columns>"},',
        '{"kind":"cell","cell":"<A1 ref>"}, or {"kind":"constant","value":<number>}.',
        'The column pattern example above is the exact formula shape the other cells follow.',
        'Choose binary/ratio/aggregate to match that shape; when it is an aggregate such as',
        '=SUM(Sheet1!$B:$B), return an aggregate IR (see the aggregate example below).',
        'Only repair cells that actually deviate from the column pattern; do not repair matching cells.',
        `Example reply: ${JSON.stringify(example)}`,
        `Aggregate example reply: ${JSON.stringify(aggregateExample)}`,
    ].join('\n');
}
function columnExample(cells) {
    return (column) => {
        const first = Object.keys(cells)
            .map((id) => {
            try {
                const parsed = parseCellId(id);
                return { id, parsed };
            }
            catch {
                return null;
            }
        })
            .filter((entry) => entry !== null && entry.parsed.sheet === column.sheet && entry.parsed.column === column.column)
            .sort((a, b) => a.parsed.row - b.parsed.row)[0];
        return `${column.sheet}!${column.column}: ${first ? cells[first.id] : 'no formula'}`;
    };
}
function stripCodeFence(text) {
    const match = /```(?:json)?\s*([\s\S]*?)```/.exec(text);
    return match ? match[1].trim() : text.trim();
}
/**
 * Wrap an LLM text function into a repair advisor: given the workbook excerpt,
 * validation anomalies, and the table schema, ask the model for IR repairs and
 * compile them into concrete CellPatches.
 */
export function createLlmRepairAdvisor(llm, table, signal) {
    return async (cells, result) => {
        const anomalies = result.anomalies.filter((anomaly) => anomaly.kind !== 'circular-reference');
        if (anomalies.length === 0)
            return [];
        const text = await llm(buildRepairPrompt(cells, anomalies, table, result.columns), signal);
        const reply = JSON.parse(stripCodeFence(text));
        if (!Array.isArray(reply.repairs)) {
            throw new Error('LLM repair reply must contain a repairs array');
        }
        const patches = [];
        for (const item of reply.repairs) {
            const resolvedId = resolveCellId(cells, item.id);
            if (!resolvedId)
                continue;
            const oldValue = cells[resolvedId]?.trim();
            if (!oldValue)
                continue;
            try {
                patches.push({
                    id: resolvedId,
                    kind: 'formula',
                    oldValue,
                    newValue: compileFormula(normalizeIr(item.ir, table), { baseCell: item.baseCell, table }),
                });
            }
            catch {
                // Skip malformed IR items instead of failing the whole repair run.
            }
        }
        return patches;
    };
}
function resolveCellId(cells, id) {
    if (cells[id])
        return id;
    const upper = id.toUpperCase();
    return Object.keys(cells).find((key) => key.split('!').pop()?.toUpperCase() === upper);
}
/**
 * Tolerate common model mistakes: bare strings as operands become cell or
 * column operands instead of failing schema validation, and an aggregate SUM
 * without a metric falls back to the table's first column.
 */
export function normalizeIr(ir, table) {
    switch (ir.operation) {
        case 'binary':
            return { ...ir, left: normalizeOperand(ir.left), right: normalizeOperand(ir.right) };
        case 'ratio':
            return { ...ir, numerator: normalizeOperand(ir.numerator), denominator: normalizeOperand(ir.denominator) };
        case 'aggregate': {
            if (!ir.metric && ir.function === 'SUM' && table) {
                const firstColumn = Object.keys(table.columns).sort((a, b) => columnToNumber(table.columns[a]) - columnToNumber(table.columns[b]))[0];
                if (firstColumn)
                    return { ...ir, metric: firstColumn };
            }
            return ir;
        }
    }
}
function normalizeOperand(operand) {
    if (typeof operand === 'string') {
        return /^[A-Za-z]{1,3}[0-9]+$/.test(operand)
            ? { kind: 'cell', cell: operand }
            : { kind: 'column', column: operand };
    }
    return operand;
}
