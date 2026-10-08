export const DEFAULT_SHEET = 'Sheet1';
export function normalizeSheet(sheet) {
    return (sheet ?? DEFAULT_SHEET).replace(/^'|'$/g, '').toUpperCase();
}
export function canonicalCellId(sheet, column, row) {
    return `${normalizeSheet(sheet)}!${column.toUpperCase()}${row}`;
}
/**
 * Convert a column letter (A … Z, AA … XFD) to its 1-based number.
 *
 * Validation belongs here rather than downstream. This used to fold *any*
 * string into a number — `columnToNumber('Sheet1')` returned 229493717 — so a
 * model that emitted a column *name* where a letter was expected sailed past
 * every guard and only failed much later, as `Invalid column letter: SHEESK`
 * (the round-trip of 229493717) or, worse, as `sheet.getColumn(229493717)`
 * allocating until the process ran out of memory. Rejecting at the boundary
 * keeps the error next to the argument that caused it.
 */
export function columnToNumber(column) {
    const normalized = column.toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(normalized))
        throw new Error(`invalid column letter: ${column}`);
    let value = 0;
    for (const char of normalized) {
        value = value * 26 + (char.charCodeAt(0) - 64);
    }
    if (value > 16384)
        throw new Error(`column out of range: ${column} (Excel supports A..XFD)`);
    return value;
}
export function numberToColumn(value) {
    let result = '';
    let n = value;
    while (n > 0) {
        const remainder = (n - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        n = Math.floor((n - 1) / 26);
    }
    return result;
}
const REF_PATTERNS = [
    /'([^']+)'!(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/y,
    /([^\s!(),+\-*/^<>=&]+)!(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/y,
    /(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/y,
    /'([^']+)'!(\$?)([A-Za-z]{1,3})(?::(\$?)([A-Za-z]{1,3}))?(?![A-Za-z0-9_(])/y,
    /([^\s!(),+\-*/^<>=&]+)!(\$?)([A-Za-z]{1,3})(?::(\$?)([A-Za-z]{1,3}))?(?![A-Za-z0-9_(])/y,
    /(\$?)([A-Za-z]{1,3})(?::(\$?)([A-Za-z]{1,3}))?(?![A-Za-z0-9_(])/y,
];
const RANGE_EXTENSION = /^:(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/;
function makePoint(sheet, columnToken, rowToken) {
    return {
        sheet: sheet === null ? null : normalizeSheet(sheet),
        column: (columnToken ?? '').replace('$', '').toUpperCase(),
        row: rowToken === null ? null : Number(rowToken.replace('$', '')),
        absColumn: columnToken?.startsWith('$') ?? false,
        absRow: rowToken?.startsWith('$') ?? false,
    };
}
function toParsedRef(match, patternIndex) {
    switch (patternIndex) {
        case 0:
        case 1:
            return { start: makePoint(match[1], match[2] + match[3], match[4] + match[5]), end: null, range: { start: 0, end: 0 } };
        case 2:
            return { start: makePoint(null, match[1] + match[2], match[3] + match[4]), end: null, range: { start: 0, end: 0 } };
        case 3:
        case 4: {
            const start = makePoint(match[1], match[2] + match[3], null);
            return { start, end: match[5] ? makePoint(match[1], match[4] + match[5], null) : null, range: { start: 0, end: 0 } };
        }
        case 5: {
            const start = makePoint(null, match[1] + match[2], null);
            return { start, end: match[4] ? makePoint(null, match[3] + match[4], null) : null, range: { start: 0, end: 0 } };
        }
        default:
            throw new Error(`unknown pattern index: ${patternIndex}`);
    }
}
export function parseFormula(input) {
    const trimmed = input.trim();
    const raw = trimmed.startsWith('=') ? trimmed.slice(1) : trimmed;
    const sanitized = raw.replace(/"(?:[^"\\]|\\.)*"/g, ' ');
    const references = [];
    let index = 0;
    while (index < sanitized.length) {
        if (index > 0 && /[0-9.]/.test(sanitized[index - 1])) {
            index += 1;
            continue;
        }
        let matched = false;
        for (let i = 0; i < REF_PATTERNS.length; i++) {
            const pattern = REF_PATTERNS[i];
            pattern.lastIndex = index;
            const match = pattern.exec(sanitized);
            if (match && match.index === index) {
                const ref = toParsedRef(match, i);
                references.push(ref);
                index = pattern.lastIndex;
                if (i <= 2) {
                    const extension = RANGE_EXTENSION.exec(sanitized.slice(index));
                    if (extension) {
                        ref.end = makePoint(ref.start.sheet, extension[1] + extension[2], extension[3] + extension[4]);
                        index += extension[0].length;
                    }
                }
                ref.range = { start: match.index, end: index };
                matched = true;
                break;
            }
        }
        if (!matched)
            index += 1;
    }
    return {
        raw,
        references: references.filter((ref) => {
            const start = ref.start;
            return !(start.row === null && ref.end === null && !start.absColumn && start.sheet === null);
        }),
    };
}
export function parseCellId(id) {
    const bang = id.lastIndexOf('!');
    const rawSheet = bang >= 0 ? id.slice(0, bang) : DEFAULT_SHEET;
    const cell = bang >= 0 ? id.slice(bang + 1) : id;
    const match = /^([A-Za-z]{1,3})(\d+)$/.exec(cell);
    if (!match) {
        throw new Error(`invalid cell id: ${id}`);
    }
    // Excel rows are 1-based and stop at 1048576. Accepting row 0 or an
    // arbitrary large row produced ids that later reached `sheet.getCell`, where
    // the same out-of-range value that blew up column handling can allocate.
    const row = Number(match[2]);
    if (row < 1 || row > 1048576) {
        throw new Error(`invalid cell id: ${id} (row out of range)`);
    }
    return { sheet: normalizeSheet(rawSheet), column: match[1].toUpperCase(), row };
}
/**
 * Shift every relative row reference in a formula by rowDelta, preserving
 * columns, absolute rows ($4), whole-column references, and sheet prefixes.
 * Returns the original formula when any shift would leave the sheet (row < 1).
 */
export function shiftFormulaRow(formula, rowDelta) {
    if (rowDelta === 0)
        return formula;
    const hasEquals = formula.trimStart().startsWith('=');
    const raw = hasEquals ? formula.trimStart().slice(1) : formula;
    const parsed = parseFormula(`=${raw}`);
    const edits = [];
    for (const ref of parsed.references) {
        const text = raw.slice(ref.range.start, ref.range.end);
        if (!ref.end) {
            edits.push({ start: ref.range.start, end: ref.range.end, text: shiftReferenceToken(text, ref.start, rowDelta) });
            continue;
        }
        const colon = text.indexOf(':');
        if (colon < 0)
            continue;
        const startToken = text.slice(0, colon);
        const endToken = text.slice(colon + 1);
        const newStart = shiftReferenceToken(startToken, ref.start, rowDelta);
        const newEnd = shiftReferenceToken(endToken, ref.end, rowDelta);
        edits.push({ start: ref.range.start, end: ref.range.end, text: `${newStart}:${newEnd}` });
    }
    let result = raw;
    for (const edit of [...edits].sort((a, b) => b.start - a.start)) {
        result = `${result.slice(0, edit.start)}${edit.text}${result.slice(edit.end)}`;
    }
    return hasEquals ? `=${result}` : result;
}
function shiftReferenceToken(token, point, rowDelta) {
    if (point.row === null || point.absRow)
        return token;
    const newRow = point.row + rowDelta;
    if (newRow < 1)
        return token;
    const match = /^(.*[A-Za-z])(\$?)(\d+)$/.exec(token);
    if (!match)
        return token;
    return `${match[1]}${newRow}`;
}
