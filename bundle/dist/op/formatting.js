import { numberToColumn } from '../formula.js';
import { parseRange } from './core.js';
export function applyDataValidation(workbook, options) {
    const parsed = parseRange(workbook, options.range);
    const validation = {
        type: options.type,
        operator: options.operator,
        formulae: [],
        allowBlank: options.allowBlank,
        showInputMessage: options.showInputMessage,
        prompt: options.prompt,
        showErrorMessage: options.showErrorMessage,
        errorStyle: options.errorStyle,
        errorTitle: options.errorTitle,
        error: options.error,
    };
    if (options.type === 'list') {
        if (!options.formula1)
            throw new Error('list data validation requires formula1 (comma-separated items or a range)');
        validation.formulae = [looksLikeRange(options.formula1) ? options.formula1 : `"${options.formula1}"`];
    }
    else if (options.formula1 !== undefined) {
        validation.formulae = [options.formula1];
        if (options.formula2 !== undefined)
            validation.formulae.push(options.formula2);
    }
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            parsed.sheet.getCell(`${numberToColumn(col)}${row}`).dataValidation = validation;
        }
    }
}
export function looksLikeRange(value) {
    return /^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$/.test(value) || /[!$]/.test(value);
}
export function applyConditionalFormatting(workbook, range, rules) {
    const parsed = parseRange(workbook, range);
    const mapped = rules.map((rule) => {
        const style = rule.style ? excelStyleToWorkbookStyle(rule.style) : undefined;
        if (rule.type === 'cellIs') {
            if (!rule.operator || rule.formula === undefined) {
                throw new Error('cellIs conditional formatting requires operator and formula');
            }
            return {
                type: 'cellIs',
                operator: rule.operator,
                formulae: [rule.formula, ...(rule.formula2 !== undefined ? [rule.formula2] : [])],
                style,
            };
        }
        if (rule.type === 'containsText') {
            if (!rule.text)
                throw new Error('containsText conditional formatting requires text');
            return {
                type: 'containsText',
                operator: 'containsText',
                text: rule.text,
                formulae: [`NOT(ISERROR(SEARCH("${rule.text}",A1)))`],
                style,
            };
        }
        if (rule.type === 'notContainsText') {
            if (!rule.text)
                throw new Error('notContainsText conditional formatting requires text');
            return {
                type: 'expression',
                formulae: [`ISERROR(SEARCH("${rule.text}",A1))`],
                style,
            };
        }
        if (rule.type === 'blanks') {
            return { type: 'expression', formulae: ['ISBLANK(A1)'], style };
        }
        if (rule.type === 'noBlanks') {
            return { type: 'expression', formulae: ['NOT(ISBLANK(A1))'], style };
        }
        if (rule.type === 'errors') {
            return { type: 'expression', formulae: ['ISERROR(A1)'], style };
        }
        if (rule.type === 'noErrors') {
            return { type: 'expression', formulae: ['NOT(ISERROR(A1))'], style };
        }
        if (rule.type === 'duplicateValues' || rule.type === 'uniqueValues') {
            const range = `$${numberToColumn(parsed.startCol)}$${parsed.startRow}:$${numberToColumn(parsed.endCol)}$${parsed.endRow}`;
            const formula = rule.type === 'duplicateValues' ? `COUNTIF(${range},A1)>1` : `COUNTIF(${range},A1)=1`;
            return { type: 'expression', formulae: [formula], style };
        }
        if (rule.type === 'aboveAverage') {
            return { type: 'aboveAverage', style };
        }
        if (rule.type === 'belowAverage') {
            return { type: 'aboveAverage', aboveAverage: false, style };
        }
        if (rule.type === 'timePeriod') {
            return { type: 'timePeriod', timePeriod: rule.timePeriod ?? 'today', style };
        }
        if (rule.type === 'dataBar') {
            return {
                type: 'dataBar',
                color: { argb: normalizeColor(rule.color ?? '638EC6') },
                cfvo: [{ type: 'min' }, { type: 'max' }],
            };
        }
        if (rule.type === 'colorScale') {
            return {
                type: 'colorScale',
                cfvo: [
                    { type: 'min' },
                    { type: 'percentile', value: 50 },
                    { type: 'max' },
                ],
                color: [
                    { argb: normalizeColor(rule.minColor ?? 'F8696B') },
                    { argb: normalizeColor(rule.midColor ?? 'FFEB84') },
                    { argb: normalizeColor(rule.maxColor ?? '63BE7B') },
                ],
            };
        }
        if (rule.type === 'iconSet') {
            return {
                type: 'iconSet',
                iconSet: rule.iconSet ?? '3Arrows',
                cfvo: [
                    { type: 'percent', value: 0 },
                    { type: 'percent', value: 33 },
                    { type: 'percent', value: 67 },
                ],
            };
        }
        if (rule.type === 'top10') {
            return { type: 'top10', rank: rule.rank ?? 10, percent: rule.percent ?? false, bottom: rule.bottom ?? false };
        }
        return { type: 'expression', formulae: [String(rule.formula ?? '')], style };
    });
    const ref = `${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${parsed.endRow}`;
    parsed.sheet.addConditionalFormatting({ ref, rules: mapped });
}
export function excelStyleToWorkbookStyle(style) {
    const result = {};
    if (style.bold !== undefined || style.italic !== undefined || style.underline !== undefined || style.fontColor !== undefined) {
        result.font = {
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            color: style.fontColor ? { argb: normalizeColor(style.fontColor) } : undefined,
        };
    }
    if (style.fill !== undefined) {
        result.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: normalizeColor(style.fill) } };
    }
    return result;
}
export function addTable(workbook, options) {
    const parsed = parseRange(workbook, options.range);
    const ref = `${numberToColumn(parsed.startCol)}${parsed.startRow}:${numberToColumn(parsed.endCol)}${parsed.endRow}`;
    const headerRow = options.headerRow ?? true;
    const header = headerRow ? parsed.startRow : null;
    const dataStart = headerRow ? parsed.startRow + 1 : parsed.startRow;
    const columns = [];
    for (let col = parsed.startCol; col <= parsed.endCol; col++) {
        const letter = numberToColumn(col);
        const nameCell = header ? parsed.sheet.getCell(`${letter}${header}`).value : null;
        columns.push({ name: nameCell === null || nameCell === undefined ? `Column${letter}` : String(nameCell) });
    }
    const rows = [];
    for (let row = dataStart; row <= parsed.endRow; row++) {
        const values = [];
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            values.push(parsed.sheet.getCell(`${numberToColumn(col)}${row}`).value);
        }
        rows.push(values);
    }
    parsed.sheet.addTable({
        name: options.name,
        ref,
        headerRow,
        totalsRow: options.totalsRow ?? false,
        columns,
        rows,
        style: {
            showRowStripes: options.showRowStripes ?? true,
            showColumnStripes: options.showColumnStripes ?? false,
        },
    });
}
export function applyStyle(workbook, range, style) {
    const parsed = parseRange(workbook, range);
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
            const cell = parsed.sheet.getCell(`${numberToColumn(col)}${row}`);
            const font = cell.font ?? {};
            if (style.bold !== undefined ||
                style.italic !== undefined ||
                style.underline !== undefined ||
                style.strikeThrough !== undefined ||
                style.fontColor !== undefined ||
                style.fontSize !== undefined ||
                style.fontName !== undefined) {
                cell.font = {
                    ...font,
                    bold: style.bold ?? font.bold,
                    italic: style.italic ?? font.italic,
                    underline: style.underline ?? font.underline,
                    strike: style.strikeThrough ?? font.strike,
                    size: style.fontSize ?? font.size,
                    name: style.fontName ?? font.name,
                    color: style.fontColor ? { argb: normalizeColor(style.fontColor) } : font.color,
                };
            }
            if (style.fill !== undefined) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: normalizeColor(style.fill) } };
            }
            if (style.numberFormat !== undefined)
                cell.numFmt = style.numberFormat;
            const alignment = cell.alignment ?? {};
            if (style.hAlign !== undefined ||
                style.vAlign !== undefined ||
                style.wrapText !== undefined ||
                style.textRotation !== undefined ||
                style.shrinkToFit !== undefined ||
                style.indent !== undefined) {
                cell.alignment = {
                    ...alignment,
                    horizontal: style.hAlign ?? alignment.horizontal,
                    vertical: style.vAlign ?? alignment.vertical,
                    wrapText: style.wrapText ?? alignment.wrapText,
                    textRotation: style.textRotation ?? alignment.textRotation,
                    shrinkToFit: style.shrinkToFit ?? alignment.shrinkToFit,
                    indent: style.indent ?? alignment.indent,
                };
            }
            if (style.border) {
                const border = {};
                for (const side of ['top', 'bottom', 'left', 'right']) {
                    const edge = style.border[side];
                    if (edge) {
                        border[side] = {
                            style: edge.style ?? 'thin',
                            color: edge.color ? { argb: normalizeColor(edge.color) } : undefined,
                        };
                    }
                }
                cell.border = border;
            }
        }
    }
}
export function normalizeColor(color) {
    const hex = color.replace('#', '').trim();
    if (/^[0-9A-Fa-f]{6}$/.test(hex))
        return `FF${hex.toUpperCase()}`;
    if (/^[0-9A-Fa-f]{8}$/.test(hex))
        return hex.toUpperCase();
    throw new Error(`invalid color: ${color} (use 6-digit hex like FF0000)`);
}
