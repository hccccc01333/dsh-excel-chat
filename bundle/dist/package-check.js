/**
 * Structural checks on a written package.
 *
 * The merge in `preserve.ts` edits relationships across four files. Getting one
 * wrong produces a package Excel refuses to open — worse than the dropped pivot
 * table it was meant to save. So the result is checked before it is allowed to
 * replace anything: if these fail, the caller refuses the write and the original
 * file stays untouched.
 *
 * This is deliberately a structural check, not a schema check. It answers "does
 * every reference resolve to something that exists", which is the class of mistake
 * hand-written package surgery actually makes. It runs anywhere Node runs, which
 * matters because CI has no Excel to open the file with.
 */
import { strFromU8, unzipSync } from 'fflate';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** Resolve a relationship target against the part that declares it. */
function resolvePart(relsPath, target) {
    // `xl/_rels/workbook.xml.rels` describes `xl/workbook.xml`, so targets are
    // relative to `xl/`. Strip the `_rels/<name>.rels` tail, then normalise.
    // `_rels/.rels` describes the package root, so its base is empty — the `*`
    // rather than `+` in the pattern is what covers that case.
    const base = relsPath.replace(/_rels\/[^/]*\.rels$/, '');
    if (target.startsWith('/'))
        return target.slice(1);
    const segments = (base + target).split('/');
    const out = [];
    for (const segment of segments) {
        if (segment === '.' || segment === '')
            continue;
        if (segment === '..')
            out.pop();
        else
            out.push(segment);
    }
    return out.join('/');
}
/**
 * Problems that would make the package unreadable. An empty array means every
 * reference in it resolves.
 */
export function validateStructure(data) {
    const problems = [];
    let files;
    try {
        files = unzipSync(data);
    }
    catch (error) {
        return [`not a readable package: ${String(error.message)}`];
    }
    const text = (name) => (files[name] ? strFromU8(files[name]) : '');
    const relsFiles = Object.keys(files).filter((name) => name.endsWith('.rels'));
    for (const relsPath of relsFiles) {
        const rels = text(relsPath);
        for (const match of rels.matchAll(/<Relationship\b[^>]*\/>/g)) {
            const xml = match[0];
            const target = /\bTarget="([^"]+)"/.exec(xml)?.[1];
            if (!target) {
                problems.push(`${relsPath}: a relationship has no Target`);
                continue;
            }
            // External targets point outside the package on purpose.
            if (/\bTargetMode="External"/.test(xml))
                continue;
            const resolved = resolvePart(relsPath, target);
            if (!files[resolved])
                problems.push(`${relsPath}: Target="${target}" resolves to ${resolved}, which is not in the package`);
        }
    }
    // Every r:id used in a part has to be declared by that part's rels.
    const idsOf = (relsPath) => new Set([...text(relsPath).matchAll(/\bId="([^"]+)"/g)].map((m) => m[1]));
    const usedIds = (xml) => [...xml.matchAll(/\br:id="([^"]+)"/g)].map((m) => m[1]);
    const workbookIds = idsOf('xl/_rels/workbook.xml.rels');
    for (const id of usedIds(text('xl/workbook.xml'))) {
        if (!workbookIds.has(id))
            problems.push(`xl/workbook.xml references ${id}, which xl/_rels/workbook.xml.rels does not declare`);
    }
    for (const sheetPath of Object.keys(files).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))) {
        const relsPath = `xl/worksheets/_rels/${sheetPath.split('/').pop()}.rels`;
        const declared = idsOf(relsPath);
        for (const id of usedIds(text(sheetPath))) {
            if (!declared.has(id))
                problems.push(`${sheetPath} references ${id}, which ${relsPath} does not declare`);
        }
    }
    const contentTypes = text('[Content_Types].xml');
    if (!contentTypes)
        problems.push('[Content_Types].xml is missing');
    for (const match of contentTypes.matchAll(/<Override\b[^>]*\/>/g)) {
        const part = /\bPartName="([^"]+)"/.exec(match[0])?.[1];
        if (part && !files[part.replace(/^\//, '')])
            problems.push(`[Content_Types].xml declares ${part}, which is not in the package`);
    }
    // A pivot part that exists but is not reachable from the workbook is invisible
    // to Excel, which is indistinguishable from having lost it.
    const pivotParts = Object.keys(files).filter((name) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(name));
    if (pivotParts.length > 0) {
        const workbookRels = text('xl/_rels/workbook.xml.rels');
        const reachable = [...workbookRels.matchAll(/<Relationship\b[^>]*\/>/g)].some((match) => /\bType="[^"]*\/pivotCacheDefinition"/.test(match[0]));
        if (!reachable)
            problems.push('the package has pivot tables but no pivotCacheDefinition relationship, so Excel cannot see them');
        if (!/<pivotCaches\b/.test(text('xl/workbook.xml')))
            problems.push('the package has pivot tables but xl/workbook.xml declares no <pivotCaches>');
    }
    void REL_NS;
    return problems;
}
