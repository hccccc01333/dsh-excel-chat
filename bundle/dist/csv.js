/**
 * RFC 4180-ish CSV parsing/writing with configurable delimiters, plus
 * formula-injection guarding for exported cells (borrowed from the
 * noatmark-dsh-plugin idea: neutralize values starting with = + - @).
 */
export function parseCsv(text, delimiter = ',') {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;
    let index = 0;
    while (index < text.length) {
        const char = text[index];
        if (inQuotes) {
            if (char === '"') {
                if (text[index + 1] === '"') {
                    field += '"';
                    index += 2;
                    continue;
                }
                inQuotes = false;
                index += 1;
                continue;
            }
            field += char;
            index += 1;
            continue;
        }
        if (char === '"' && field === '') {
            inQuotes = true;
            index += 1;
            continue;
        }
        if (char === delimiter) {
            row.push(field);
            field = '';
            index += 1;
            continue;
        }
        if (char === '\n' || char === '\r') {
            if (char === '\r' && text[index + 1] === '\n')
                index += 1;
            row.push(field);
            field = '';
            rows.push(row);
            row = [];
            index += 1;
            continue;
        }
        field += char;
        index += 1;
    }
    row.push(field);
    rows.push(row);
    return rows;
}
function csvField(value, delimiter) {
    if (value.includes(delimiter) || value.includes('"') || value.includes('\n') || value.includes('\r')) {
        return `"${value.replaceAll('"', '""')}"`;
    }
    return value;
}
/** Neutralize spreadsheet formula injection (=, +, -, @) for literal values. */
export function guardFormulaInjection(value) {
    return /^[=+\-@]/.test(value) ? `'${value}` : value;
}
export function stringifyCsv(rows, delimiter = ',') {
    return rows.map((row) => row.map((cell) => csvField(cell, delimiter)).join(delimiter)).join('\r\n') + '\r\n';
}
