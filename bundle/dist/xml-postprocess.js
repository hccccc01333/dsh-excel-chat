import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
export function emptyAnnotations() {
    return { comments: new Map(), sparklines: new Map() };
}
function escapeXml(text) {
    return text
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;');
}
/**
 * Rewrite the saved xlsx zip: inject comments parts, VML shapes, sparkline
 * extensions, and the worksheet plumbing (legacyDrawing + rels + content
 * types) they require.
 */
export function annotateWorkbookXml(data, annotations, sheetFileOf) {
    if (annotations.comments.size === 0 && annotations.sparklines.size === 0)
        return data;
    const files = unzipSync(data);
    // ExcelJS names sheets xl/worksheets/sheetN.xml in id order; map sheet
    // names to their file via the workbook.xml sheet list.
    const workbookXml = strFromU8(files['xl/workbook.xml'] ?? new Uint8Array(0));
    const relsXml = strFromU8(files['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(0));
    const ridTarget = new Map();
    for (const match of relsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
        ridTarget.set(match[1], match[2]);
    }
    for (const match of workbookXml.matchAll(/<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]+)"/g)) {
        const name = match[1].replaceAll('&amp;', '&').replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&quot;', '"').replaceAll('&#39;', "'");
        const target = ridTarget.get(match[2]);
        if (target)
            sheetFileOf.set(name, `xl/${target.replace(/^\/?xl\//, '').replace(/^\//, '')}`);
    }
    let contentTypes = strFromU8(files['[Content_Types].xml'] ?? new Uint8Array(0));
    let commentFileIndex = 0;
    let vmlFileIndex = 0;
    for (const [sheetName, comments] of annotations.comments) {
        const sheetFile = sheetFileOf.get(sheetName);
        if (!sheetFile)
            throw new Error(`sheet not found for comments: ${sheetName}`);
        commentFileIndex += 1;
        const commentsFile = `xl/comments${commentFileIndex}.xml`;
        const author = comments[0].author;
        files[commentsFile] = strToU8(commentsXml(author, comments));
        contentTypes = addOverride(contentTypes, `/${commentsFile}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
        if (!/<Default Extension="vml"/.test(contentTypes)) {
            contentTypes = contentTypes.replace(/(<Types[^>]*>)/, '$1<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vml"/>');
        }
        vmlFileIndex += 1;
        const vmlFile = `xl/drawings/vmlDrawing${vmlFileIndex}.vml`;
        files[vmlFile] = strToU8(commentVml(comments));
        patchSheetForComments(files, sheetFile, commentsFile, vmlFile);
    }
    for (const [sheetName, groups] of annotations.sparklines) {
        const sheetFile = sheetFileOf.get(sheetName);
        if (!sheetFile)
            throw new Error(`sheet not found for sparklines: ${sheetName}`);
        files[sheetFile] = strToU8(addSparklineExt(strFromU8(files[sheetFile] ?? new Uint8Array(0)), groups));
    }
    files['[Content_Types].xml'] = strToU8(contentTypes);
    return Buffer.from(zipSync(files));
}
function commentsXml(author, comments) {
    const items = comments.map((comment) => `<comment ref="${escapeXml(comment.ref)}" authorId="0"><text><r><t xml:space="preserve">${escapeXml(comment.text)}</t></r></text></comment>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>${escapeXml(author)}</author></authors><commentList>${items}</commentList></comments>`;
}
function commentVml(comments) {
    const shapes = comments.map((comment, index) => {
        const shapeId = 1025 + index;
        return `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202" style="position:absolute;margin-left:0pt;margin-top:0pt;width:${comment.width}pt;height:${comment.height}pt;z-index:${index + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>${rowOf(comment.ref) - 1}</x:Row><x:Column>${columnIndexOf(comment.ref)}</x:Column></x:ClientData></v:shape>`;
    }).join('');
    return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:fill on="f" focussize="0,0"/><v:stroke on="f"/><v:path gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>${shapes}</xml>`;
}
function rowOf(ref) {
    const match = /^([A-Za-z]{1,3})(\d+)$/.exec(ref);
    if (!match)
        throw new Error(`invalid cell reference: ${ref}`);
    return Number(match[2]);
}
function columnIndexOf(ref) {
    const match = /^([A-Za-z]{1,3})(\d+)$/.exec(ref);
    if (!match)
        throw new Error(`invalid cell reference: ${ref}`);
    let index = 0;
    for (const char of match[1].toUpperCase()) {
        index = index * 26 + (char.charCodeAt(0) - 64);
    }
    return index - 1;
}
/** Append the legacyDrawing plumbing a sheet needs before comments render. */
function patchSheetForComments(files, sheetFile, commentsFile, vmlFile) {
    const sheetPath = sheetFile;
    const relsPath = `${sheetPath.replace('xl/worksheets/', 'xl/worksheets/_rels/')}.rels`;
    const commentsRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
    const vmlRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
    let rels = files[relsPath] ? strFromU8(files[relsPath]) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    let nextRid = 1;
    for (const match of rels.matchAll(/Id="rId(\d+)"/g))
        nextRid = Math.max(nextRid, Number(match[1]) + 1);
    const commentsRid = `rId${nextRid}`;
    const vmlRid = `rId${nextRid + 1}`;
    rels = rels.replace('</Relationships>', `<Relationship Id="${commentsRid}" Type="${commentsRelType}" Target="../${commentsFile.replace('xl/', '')}"/><Relationship Id="${vmlRid}" Type="${vmlRelType}" Target="../${vmlFile.replace('xl/', '')}"/></Relationships>`);
    files[relsPath] = strToU8(rels);
    const xml = strFromU8(files[sheetPath] ?? new Uint8Array(0));
    const legacy = `<legacyDrawing r:id="${vmlRid}"/>`;
    let patched = xml;
    if (!xml.includes(legacy)) {
        // legacyDrawing must come after <drawing> and before tableParts/extLst.
        if (/<extLst/.test(xml))
            patched = xml.replace(/<extLst/, `${legacy}<extLst`);
        else if (/<tableParts/.test(xml))
            patched = xml.replace(/<tableParts/, `${legacy}<tableParts`);
        else
            patched = xml.replace('</worksheet>', `${legacy}</worksheet>`);
    }
    files[sheetPath] = strToU8(patched);
}
const SPARKLINE_EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
function addSparklineExt(xml, groups) {
    const ext = sparklineExtXml(groups);
    if (/<extLst[\s\S]*<\/extLst>/.test(xml)) {
        // Merge into the existing extension list.
        return xml.replace('</extLst>', `${ext}</extLst>`);
    }
    return xml.replace('</worksheet>', `<extLst>${ext}</extLst></worksheet>`);
}
function sparklineExtXml(groups) {
    const groupXml = groups.map((group) => {
        const rows = sparklineRows(group);
        const sparklines = rows.map(({ data, location }) => `<x14:sparkline><xm:f>${escapeXml(data)}</xm:f><xm:sqref>${escapeXml(location)}</xm:sqref></x14:sparkline>`).join('');
        const flags = [
            `type="${group.type === 'line' ? 'line' : group.type === 'column' ? 'column' : 'stacked'}"`,
            group.markers ? 'markers="1" high="1" low="1"' : '',
        ].filter(Boolean).join(' ');
        return `<x14:sparklineGroup displayEmptyCellsAs="gap" ${flags}>` +
            `<x14:colorSeries rgb="${group.color}"/>` +
            `<x14:colorNegative rgb="${group.negativeColor}"/>` +
            `<x14:colorAxis rgb="FF000000"/>` +
            `<x14:colorMarkers rgb="${group.color}"/>` +
            `<x14:colorFirst rgb="${group.color}"/>` +
            `<x14:colorLast rgb="${group.color}"/>` +
            `<x14:colorHigh rgb="${group.highColor}"/>` +
            `<x14:colorLow rgb="${group.lowColor}"/>` +
            `<x14:sparklines>${sparklines}</x14:sparklines>` +
            `</x14:sparklineGroup>`;
    }).join('');
    return `<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="${SPARKLINE_EXT_URI}"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${groupXml}</x14:sparklineGroups></ext>`;
}
/**
 * Pair each data row with its location cell: "订单!B2:F31" + "订单!G2:G31"
 * produces per-row sparklines ("订单!B2:F2" -> G2); a single-row data range
 * maps to a single location. The sparkline formula keeps the sheet name.
 */
function sparklineRows(group) {
    const dataSheet = sheetOfRange(group.dataRange);
    const locationSheet = sheetOfRange(group.locationRange);
    if (dataSheet !== locationSheet) {
        throw new Error('sparkline dataRange and locationRange must be on the same sheet');
    }
    const data = parseRangeParts(group.dataRange);
    const location = parseRangeParts(group.locationRange);
    const dataRows = data.endRow - data.startRow;
    const locationRows = location.endRow - location.startRow;
    if (dataRows !== locationRows) {
        throw new Error(`sparkline dataRange rows (${dataRows + 1}) must match locationRange rows (${locationRows + 1})`);
    }
    const out = [];
    for (let offset = 0; offset <= dataRows; offset++) {
        const row = data.startRow + offset;
        out.push({
            data: `${dataSheet}!${columnName(data.startCol + 1)}${row}:${columnName(data.endCol + 1)}${row}`,
            location: `${columnName(location.startCol + 1)}${row}`,
        });
    }
    return out;
}
function sheetOfRange(range) {
    const bang = range.lastIndexOf('!');
    if (bang < 0)
        throw new Error(`sparkline range requires a sheet: ${range}`);
    return range.slice(0, bang);
}
function parseRangeParts(range) {
    const body = range.slice(range.lastIndexOf('!') + 1);
    const match = /^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/.exec(body);
    if (!match)
        throw new Error(`invalid sparkline range: ${range}`);
    return {
        startCol: columnIndexOf(`${match[1]}1`),
        startRow: Number(match[2]),
        endCol: columnIndexOf(`${match[3]}1`),
        endRow: Number(match[4]),
    };
}
function columnName(index) {
    let name = '';
    let value = index;
    while (value > 0) {
        const remaider = (value - 1) % 26;
        name = `${String.fromCharCode(65 + remaider)}${name}`;
        value = Math.floor((value - remaider - 1) / 26);
    }
    return name;
}
function addOverride(contentTypes, partName, contentType) {
    if (contentTypes.includes(`PartName="${partName}"`))
        return contentTypes;
    return contentTypes.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`);
}
