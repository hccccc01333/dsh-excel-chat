/**
 * Carry content through an edit that ExcelJS cannot round-trip.
 *
 * ExcelJS models cells, styles and a few objects. It does not model pivot tables,
 * so a rewrite drops `xl/pivotTables/*` and `xl/pivotCache/*` along with every
 * reference to them, and the pivot table is simply gone from the saved file. The
 * parts are kept here and put back afterwards.
 *
 * Two rules make this safe to run on somebody's real workbook:
 *
 * 1. **Nothing is regenerated.** The parts are copied byte for byte from the
 *    original, and the relationships are copied verbatim, so anything already
 *    working inside them keeps working.
 * 2. **Ids are re-allocated, never reused.** The original's `rId1` means nothing in
 *    a file ExcelJS has just written, and reusing it would silently point at
 *    whatever ExcelJS happened to put there. Every injected relationship gets a
 *    fresh id and every reference to it is rewritten to match.
 *
 * A pivot table is anchored two ways, and both have to survive: the worksheet
 * relationship (`worksheets/_rels/sheetN.xml.rels`) plus, when the file uses it,
 * the worksheet's own `<tableParts>` element. Files produced by Excel tend to use
 * only the former; files that have been through other tools often have both.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
const PIVOT_PART = /^xl\/(pivotTables|pivotCache)\//;
const PIVOT_CACHE_REL = /\/pivotCacheDefinition$/;
const PIVOT_TABLE_REL = /\/pivotTable$/;
/** True when there is nothing worth carrying through an edit. */
export function isEmpty(preserved) {
    return Object.keys(preserved.parts).length === 0;
}
/** Read everything an edit would otherwise destroy out of the original file. */
export function capturePreserved(data) {
    const empty = { parts: {}, sheetRelationships: new Map(), sheetTableParts: new Map(), contentTypes: [] };
    let files;
    try {
        files = unzipSync(data);
    }
    catch {
        return empty;
    }
    const text = (name) => (files[name] ? strFromU8(files[name]) : '');
    const parts = {};
    for (const [name, bytes] of Object.entries(files)) {
        if (PIVOT_PART.test(name))
            parts[name] = bytes;
    }
    if (Object.keys(parts).length === 0)
        return empty;
    const workbook = text('xl/workbook.xml');
    const workbookRels = text('xl/_rels/workbook.xml.rels');
    const contentTypes = text('[Content_Types].xml');
    // Sheet name → part path, so anchors can be attached to the right sheet after
    // ExcelJS has renumbered the package.
    const sheetPathByName = new Map();
    const relTargetById = new Map();
    for (const match of workbookRels.matchAll(/<Relationship\b[^>]*\/>/g)) {
        const id = /\bId="([^"]+)"/.exec(match[0])?.[1];
        const target = /\bTarget="([^"]+)"/.exec(match[0])?.[1];
        if (id && target)
            relTargetById.set(id, target);
    }
    for (const match of workbook.matchAll(/<sheet\b[^>]*\/>/g)) {
        const name = /\bname="([^"]*)"/.exec(match[0])?.[1];
        const rid = /\br:id="([^"]+)"/.exec(match[0])?.[1];
        const target = rid ? relTargetById.get(rid) : undefined;
        if (name && target)
            sheetPathByName.set(decodeXml(name), `xl/${target.replace(/^\/?xl\//, '')}`);
    }
    const sheetRelationships = new Map();
    const sheetTableParts = new Map();
    for (const [name, path] of sheetPathByName) {
        const relsPath = `xl/worksheets/_rels/${path.split('/').pop()}.rels`;
        const rels = text(relsPath);
        const declared = [...rels.matchAll(/<Relationship\b[^>]*\/>/g)]
            .map((match) => match[0])
            .filter((xml) => PIVOT_TABLE_REL.test(/\bType="([^"]*)"/.exec(xml)?.[1] ?? ''));
        if (declared.length > 0)
            sheetRelationships.set(name, declared);
        const tableParts = /<tableParts\b[\s\S]*?<\/tableParts>/.exec(text(path));
        if (tableParts)
            sheetTableParts.set(name, tableParts[0]);
    }
    const cacheRelationship = [...workbookRels.matchAll(/<Relationship\b[^>]*\/>/g)]
        .map((match) => match[0])
        .find((xml) => PIVOT_CACHE_REL.test(/\bType="([^"]*)"/.exec(xml)?.[1] ?? ''));
    const overrides = [...contentTypes.matchAll(/<Override\b[^>]*\/>/g)]
        .map((match) => match[0])
        .filter((xml) => {
        const part = /\bPartName="([^"]+)"/.exec(xml)?.[1] ?? '';
        return PIVOT_PART.test(part.replace(/^\//, ''));
    });
    return {
        parts,
        pivotCaches: /<pivotCaches\b[\s\S]*?<\/pivotCaches>/.exec(workbook)?.[0],
        cacheRelationship,
        sheetRelationships,
        sheetTableParts,
        contentTypes: overrides,
    };
}
/**
 * Put the captured content back into a freshly written package.
 *
 * Exported for testing: it is pure, so the merge can be checked against real
 * packages without going anywhere near a file.
 */
export function applyPreserved(output, preserved) {
    if (isEmpty(preserved))
        return output;
    const files = { ...unzipSync(output) };
    const text = (name) => (files[name] ? strFromU8(files[name]) : '');
    for (const [name, bytes] of Object.entries(preserved.parts))
        files[name] = bytes;
    // Ids already in use, so a fresh one never collides with what ExcelJS wrote.
    const usedIds = (xml) => new Set([...xml.matchAll(/\bId="([^"]+)"/g)].map((m) => m[1]));
    const nextId = (used) => {
        for (let i = 1;; i++)
            if (!used.has(`rId${i}`))
                return `rId${i}`;
    };
    // 1. Workbook-level cache declaration + the relationship it points at.
    const workbookPath = 'xl/workbook.xml';
    const relsPath = 'xl/_rels/workbook.xml.rels';
    if (preserved.pivotCaches && preserved.cacheRelationship) {
        const rels = text(relsPath);
        const used = usedIds(rels);
        const id = nextId(used);
        const relationship = preserved.cacheRelationship.replace(/\bId="[^"]*"/, `Id="${id}"`);
        files[relsPath] = strToU8(rels.replace('</Relationships>', `${relationship}</Relationships>`));
        const workbook = text(workbookPath);
        const caches = preserved.pivotCaches.replace(/\br:id="[^"]*"/, `r:id="${id}"`);
        // `<pivotCaches>` belongs after `<calcPr>` and before `<extLst>` per the schema.
        const anchored = /<extLst\b/.test(workbook)
            ? workbook.replace(/<extLst\b/, `${caches}<extLst`)
            : workbook.replace('</workbook>', `${caches}</workbook>`);
        files[workbookPath] = strToU8(anchored);
    }
    // 2. Per-sheet anchors. Sheet files are renumbered by ExcelJS, so resolve the
    //    name again rather than assuming the original path still applies.
    const workbook = text(workbookPath);
    const rels = text(relsPath);
    const relTargetById = new Map();
    for (const match of rels.matchAll(/<Relationship\b[^>]*\/>/g)) {
        const id = /\bId="([^"]+)"/.exec(match[0])?.[1];
        const target = /\bTarget="([^"]+)"/.exec(match[0])?.[1];
        if (id && target)
            relTargetById.set(id, target);
    }
    for (const match of workbook.matchAll(/<sheet\b[^>]*\/>/g)) {
        const name = /\bname="([^"]*)"/.exec(match[0])?.[1];
        const rid = /\br:id="([^"]+)"/.exec(match[0])?.[1];
        if (!name)
            continue;
        const sheetName = decodeXml(name);
        const declared = preserved.sheetRelationships.get(sheetName);
        const tableParts = preserved.sheetTableParts.get(sheetName);
        if (!declared && !tableParts)
            continue;
        const target = rid ? relTargetById.get(rid) : undefined;
        if (!target)
            continue;
        const sheetPath = `xl/${target.replace(/^\/?xl\//, '')}`;
        const sheetRelsPath = `xl/worksheets/_rels/${sheetPath.split('/').pop()}.rels`;
        let sheetRels = text(sheetRelsPath) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
        const used = usedIds(sheetRels);
        const added = [];
        for (const relationship of declared ?? []) {
            const id = nextId(used);
            used.add(id);
            added.push(relationship.replace(/\bId="[^"]*"/, `Id="${id}"`));
        }
        files[sheetRelsPath] = strToU8(sheetRels.replace('</Relationships>', `${added.join('')}</Relationships>`));
        if (tableParts) {
            // The element has to name one of the relationships just added, and it has to
            // sit in the right place in the worksheet schema — before `</worksheet>`,
            // after everything else the sheet declares.
            const firstId = /\bId="([^"]+)"/.exec(added[0] ?? '')?.[1];
            const restored = firstId
                ? tableParts.replace(/<tableParts\b[^>]*>/, `<tableParts count="1"><tablePart r:id="${firstId}"/>`)
                : tableParts;
            const sheet = text(sheetPath);
            if (sheet && !/<tableParts\b/.test(sheet)) {
                files[sheetPath] = strToU8(sheet.replace('</worksheet>', `${restored}</worksheet>`));
            }
        }
    }
    // 3. Content types, or the package is unreadable however correct the rest is.
    if (preserved.contentTypes.length > 0) {
        const path = '[Content_Types].xml';
        const xml = text(path);
        const missing = preserved.contentTypes.filter((override) => {
            const part = /\bPartName="([^"]+)"/.exec(override)?.[1];
            return part !== undefined && !xml.includes(`PartName="${part}"`);
        });
        if (missing.length > 0)
            files[path] = strToU8(xml.replace('</Types>', `${missing.join('')}</Types>`));
    }
    return new Uint8Array(zipSync(files));
}
const decodeXml = (value) => value.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
