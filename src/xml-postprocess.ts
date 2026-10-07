import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

/**
 * Post-save xlsx XML annotations for features ExcelJS cannot write:
 * cell comments (comments XML + VML shapes), and sparklines (x14 ext).
 *
 * ExcelJS reads comments but silently drops them on save, so comments added
 * by `addComment` are re-injected into the saved zip here. Sparklines are a
 * Microsoft x14 worksheet extension that ExcelJS never models.
 */

export interface CommentSpec {
  /** 1-based cell reference inside the sheet, e.g. "B2". */
  ref: string
  text: string
  author: string
  /** Comment box size in points (defaults 108 x 60). */
  width: number
  height: number
}

export interface SparklineGroupSpec {
  /** Workbook-qualified data range, e.g. "订单!B2:F31". */
  dataRange: string
  /** Workbook-qualified location range, e.g. "订单!G2:G31". */
  locationRange: string
  type: 'line' | 'column' | 'stacked'
  color: string
  negativeColor: string
  markers: boolean
  highColor: string
  lowColor: string
}

export interface WorkbookAnnotations {
  comments: Map<string, CommentSpec[]>
  sparklines: Map<string, SparklineGroupSpec[]>
  /** Sheets whose saved view should show formulas instead of their results. */
  showFormulas: Set<string>
}

export function emptyAnnotations(): WorkbookAnnotations {
  return { comments: new Map(), sparklines: new Map(), showFormulas: new Set() }
}

function escapeXml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/** The inverse of {@link escapeXml}, for values read back out of a part. */
function unescapeXml(text: string): string {
  return text
    .replaceAll('&amp;', '&')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&#39;', "'")
}

/**
 * One attribute of an XML start tag, found wherever it sits in the tag.
 *
 * Reading attributes by position — `/<Relationship[^>]*Id="…"[^>]*Target="…"/` —
 * only works while every writer emits them in the same order, and this module
 * reads parts other tools produced. Looking each attribute up on its own costs
 * nothing and cannot be caught out by a reordered tag.
 */
function attributeOf(tag: string, name: string): string | null {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const match = new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag)
  if (!match) return null
  return match[1] ?? match[2] ?? null
}

/**
 * A sheet's own `<sheetView>` decides whether formulas or their results are
 * shown. exceljs' SheetViewXform renders a fixed attribute list that omits
 * showFormulas, so assigning `view.showFormulas` is silently dropped — the
 * attribute has to be injected into the saved XML instead.
 */
function withShowFormulas(sheetXml: string): string {
  // `showFormulas` belongs to a view, and a sheet can hold more than one, so
  // every view that does not already ask for it gets the attribute. (`\b` keeps
  // this from matching `<sheetViews>`.)
  if (/<sheetView\b/.test(sheetXml)) {
    return sheetXml.replace(/<sheetView\b[^>]*?\/?>/g, (tag) =>
      /\bshowFormulas\s*=/.test(tag) ? tag : tag.replace(/\/?>$/, (tail) => ` showFormulas="1"${tail}`))
  }
  // An empty `<sheetViews>` is still a list; adding a second one would be
  // invalid, so the view goes inside the list that is already there.
  const emptyList = /<sheetViews\b[^>]*>\s*<\/sheetViews>/.exec(sheetXml)
  if (emptyList) {
    const at = emptyList.index + emptyList[0].length - '</sheetViews>'.length
    return sheetXml.slice(0, at) + '<sheetView workbookViewId="0" showFormulas="1"/>' + sheetXml.slice(at)
  }
  if (/<sheetViews\b/.test(sheetXml)) return sheetXml
  // exceljs only emits <sheetViews> when the sheet already has view settings, so
  // for a plain sheet the whole block has to be inserted — in schema order,
  // which puts sheetViews right after dimension.
  const block = '<sheetViews><sheetView workbookViewId="0" showFormulas="1"/></sheetViews>'
  const anchor = /<dimension\b[^>]*?\/?>/.exec(sheetXml) ?? /<worksheet\b[^>]*?>/.exec(sheetXml)
  if (!anchor) return sheetXml
  const at = anchor.index + anchor[0].length
  return sheetXml.slice(0, at) + block + sheetXml.slice(at)
}

/**
 * Rewrite the saved xlsx zip: inject comments parts, VML shapes, sparkline
 * extensions, and the worksheet plumbing (legacyDrawing + rels + content
 * types) they require.
 */
export function annotateWorkbookXml(data: Uint8Array, annotations: WorkbookAnnotations, sheetFileOf: Map<string, string>): Uint8Array {
  if (annotations.comments.size === 0 && annotations.sparklines.size === 0 && annotations.showFormulas.size === 0) return data
  const files = unzipSync(data)

  // ExcelJS names sheets xl/worksheets/sheetN.xml in id order; map sheet
  // names to their file via the workbook.xml sheet list.
  const workbookXml = strFromU8(files['xl/workbook.xml'] ?? new Uint8Array(0))
  const relsXml = strFromU8(files['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(0))
  const ridTarget = new Map<string, string>()
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*?\/?>/g)) {
    const id = attributeOf(match[0], 'Id')
    const target = attributeOf(match[0], 'Target')
    if (id !== null && target !== null) ridTarget.set(id, target)
  }
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*?\/?>/g)) {
    const rawName = attributeOf(match[0], 'name')
    const rid = attributeOf(match[0], 'r:id')
    if (rawName === null || rid === null) continue
    const target = ridTarget.get(rid)
    if (target) sheetFileOf.set(unescapeXml(rawName), `xl/${target.replace(/^\/?xl\//, '').replace(/^\//, '')}`)
  }

  let contentTypes = strFromU8(files['[Content_Types].xml'] ?? new Uint8Array(0))
  let commentFileIndex = 0
  let vmlFileIndex = 0

  for (const [sheetName, comments] of annotations.comments) {
    const sheetFile = sheetFileOf.get(sheetName)
    if (!sheetFile) throw new Error(`sheet not found for comments: ${sheetName}`)
    commentFileIndex += 1
    const commentsFile = `xl/comments${commentFileIndex}.xml`
    files[commentsFile] = strToU8(commentsXml(comments))
    contentTypes = addOverride(contentTypes, `/${commentsFile}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml')
    if (!/<Default Extension="vml"/.test(contentTypes)) {
      contentTypes = contentTypes.replace(
        /(<Types[^>]*>)/,
        '$1<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
      )
    }
    vmlFileIndex += 1
    const vmlFile = `xl/drawings/vmlDrawing${vmlFileIndex}.vml`
    files[vmlFile] = strToU8(commentVml(comments))
    patchSheetForComments(files, sheetFile, commentsFile, vmlFile)
  }

  for (const [sheetName, groups] of annotations.sparklines) {
    const sheetFile = sheetFileOf.get(sheetName)
    if (!sheetFile) throw new Error(`sheet not found for sparklines: ${sheetName}`)
    files[sheetFile] = strToU8(addSparklineExt(strFromU8(files[sheetFile] ?? new Uint8Array(0)), groups))
  }

  for (const sheetName of annotations.showFormulas) {
    const sheetFile = sheetFileOf.get(sheetName)
    if (!sheetFile) throw new Error(`sheet not found for showFormulas: ${sheetName}`)
    files[sheetFile] = strToU8(withShowFormulas(strFromU8(files[sheetFile] ?? new Uint8Array(0))))
  }

  files['[Content_Types].xml'] = strToU8(contentTypes)
  return Buffer.from(zipSync(files))
}

function commentsXml(comments: CommentSpec[]): string {
  const authors: string[] = []
  const authorId = new Map<string, number>()
  for (const comment of comments) {
    if (!authorId.has(comment.author)) {
      authorId.set(comment.author, authors.length)
      authors.push(comment.author)
    }
  }
  const authorsXml = authors.map((author) => `<author>${escapeXml(author)}</author>`).join('')
  const items = comments.map((comment) =>
    `<comment ref="${escapeXml(comment.ref)}" authorId="${authorId.get(comment.author)}"><text><r><t xml:space="preserve">${escapeXml(comment.text)}</t></r></text></comment>`,
  ).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>${authorsXml}</authors><commentList>${items}</commentList></comments>`
}

function commentVml(comments: CommentSpec[]): string {
  const shapes = comments.map((comment, index) => {
    const shapeId = 1025 + index
    return `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202" style="position:absolute;margin-left:0pt;margin-top:0pt;width:${comment.width}pt;height:${comment.height}pt;z-index:${index + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>${rowOf(comment.ref) - 1}</x:Row><x:Column>${columnIndexOf(comment.ref)}</x:Column></x:ClientData></v:shape>`
  }).join('')
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:fill on="f" focussize="0,0"/><v:stroke on="f"/><v:path gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>${shapes}</xml>`
}

function rowOf(ref: string): number {
  const match = /^([A-Za-z]{1,3})(\d+)$/.exec(ref)
  if (!match) throw new Error(`invalid cell reference: ${ref}`)
  return Number(match[2]!)
}

function columnIndexOf(ref: string): number {
  const match = /^([A-Za-z]{1,3})(\d+)$/.exec(ref)
  if (!match) throw new Error(`invalid cell reference: ${ref}`)
  let index = 0
  for (const char of match[1]!.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64)
  }
  return index - 1
}

/** Append the legacyDrawing plumbing a sheet needs before comments render. */
function patchSheetForComments(files: Record<string, Uint8Array>, sheetFile: string, commentsFile: string, vmlFile: string): void {
  const sheetPath = sheetFile
  const relsPath = `${sheetPath.replace('xl/worksheets/', 'xl/worksheets/_rels/')}.rels`
  const commentsRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
  const vmlRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'

  let rels = files[relsPath] ? strFromU8(files[relsPath]) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let nextRid = 1
  for (const match of rels.matchAll(/Id="rId(\d+)"/g)) nextRid = Math.max(nextRid, Number(match[1]) + 1)
  const commentsRid = `rId${nextRid}`
  const vmlRid = `rId${nextRid + 1}`
  rels = rels.replace('</Relationships>', `<Relationship Id="${commentsRid}" Type="${commentsRelType}" Target="../${commentsFile.replace('xl/', '')}"/><Relationship Id="${vmlRid}" Type="${vmlRelType}" Target="../${vmlFile.replace('xl/', '')}"/></Relationships>`)
  files[relsPath] = strToU8(rels)

  const xml = strFromU8(files[sheetPath] ?? new Uint8Array(0))
  const legacy = `<legacyDrawing r:id="${vmlRid}"/>`
  let patched = xml
  if (!xml.includes(legacy)) {
    // legacyDrawing sits after <drawing> and before tableParts, which itself
    // precedes extLst. Inserting at whichever of those two comes first keeps
    // that order. Matching the first `<extLst` instead would drop it inside a
    // nested list — the one a conditional-formatting rule carries — and make
    // the sheet invalid.
    const bounds = [worksheetExtListStart(xml), xml.indexOf('<tableParts')].filter((at) => at >= 0)
    const at = bounds.length > 0 ? Math.min(...bounds) : -1
    patched = at >= 0
      ? xml.slice(0, at) + legacy + xml.slice(at)
      : xml.replace('</worksheet>', `${legacy}</worksheet>`)
  }
  files[sheetPath] = strToU8(patched)
}

const SPARKLINE_EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'

/**
 * Index of the worksheet-level `<extLst>`'s opening tag, or -1 when the sheet
 * has none.
 *
 * Extensions nest: a conditional-formatting rule carries a list of its own, and
 * exceljs writes one for every data bar or colour scale. So neither "the first
 * `<extLst>`" nor "the first closing tag" identifies the worksheet's own list —
 * both resolve inside that rule, where an appended extension is ignored by Excel
 * and a sibling inserted there makes the sheet invalid. Walking the tags and
 * taking the list that sits directly under `<worksheet>` is what identifies it.
 */
function worksheetExtListStart(xml: string): number {
  const tag = /<(\/?)([A-Za-z_][\w:.-]*)((?:"[^"]*"|[^>"])*?)(\/?)>/g
  let depth = 0
  for (const match of xml.matchAll(tag)) {
    const [full, closing, , , selfClosing] = match
    if (closing === '/') {
      depth -= 1
      continue
    }
    if (depth === 1 && selfClosing !== '/' && full.startsWith('<extLst')) return match.index!
    if (selfClosing !== '/') depth += 1
  }
  return -1
}

function addSparklineExt(xml: string, groups: SparklineGroupSpec[]): string {
  const ext = sparklineExtXml(groups)
  if (worksheetExtListStart(xml) >= 0) {
    // The worksheet's own list is its last child, so its closing tag is the last
    // one in the document; nested lists close earlier.
    const at = xml.lastIndexOf('</extLst>')
    return xml.slice(0, at) + ext + xml.slice(at)
  }
  return xml.replace('</worksheet>', `<extLst>${ext}</extLst></worksheet>`)
}

function sparklineExtXml(groups: SparklineGroupSpec[]): string {
  const groupXml = groups.map((group) => {
    const rows = sparklineRows(group)
    const sparklines = rows.map(({ data, location }) =>
      `<x14:sparkline><xm:f>${escapeXml(data)}</xm:f><xm:sqref>${escapeXml(location)}</xm:sqref></x14:sparkline>`,
    ).join('')
    const flags = [
      `type="${group.type === 'line' ? 'line' : group.type === 'column' ? 'column' : 'stacked'}"`,
      group.markers ? 'markers="1" high="1" low="1"' : '',
    ].filter(Boolean).join(' ')
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
      `</x14:sparklineGroup>`
  }).join('')
  return `<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="${SPARKLINE_EXT_URI}"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${groupXml}</x14:sparklineGroups></ext>`
}

/**
 * Pair each data row with its location cell: "订单!B2:F31" + "订单!G2:G31"
 * produces per-row sparklines ("订单!B2:F2" -> G2); a single-row data range
 * maps to a single location. The sparkline formula keeps the sheet name.
 */
function sparklineRows(group: SparklineGroupSpec): Array<{ data: string; location: string }> {
  const dataSheet = sheetOfRange(group.dataRange)
  const locationSheet = sheetOfRange(group.locationRange)
  if (dataSheet !== locationSheet) {
    throw new Error('sparkline dataRange and locationRange must be on the same sheet')
  }
  const data = parseRangeParts(group.dataRange)
  const location = parseRangeParts(group.locationRange)
  const dataRows = data.endRow - data.startRow
  const locationRows = location.endRow - location.startRow
  if (dataRows !== locationRows) {
    throw new Error(`sparkline dataRange rows (${dataRows + 1}) must match locationRange rows (${locationRows + 1})`)
  }
  const out: Array<{ data: string; location: string }> = []
  for (let offset = 0; offset <= dataRows; offset++) {
    const dataRow = data.startRow + offset
    const locationRow = location.startRow + offset
    out.push({
      data: `${dataSheet}!${columnName(data.startCol + 1)}${dataRow}:${columnName(data.endCol + 1)}${dataRow}`,
      location: `${columnName(location.startCol + 1)}${locationRow}`,
    })
  }
  return out
}

function sheetOfRange(range: string): string {
  const bang = range.lastIndexOf('!')
  if (bang < 0) throw new Error(`sparkline range requires a sheet: ${range}`)
  return range.slice(0, bang)
}

function parseRangeParts(range: string): { startCol: number; startRow: number; endCol: number; endRow: number } {
  const body = range.slice(range.lastIndexOf('!') + 1)
  const match = /^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/.exec(body)
  if (!match) throw new Error(`invalid sparkline range: ${range}`)
  return {
    startCol: columnIndexOf(`${match[1]}1`),
    startRow: Number(match[2]),
    endCol: columnIndexOf(`${match[3]}1`),
    endRow: Number(match[4]),
  }
}

function columnName(index: number): string {
  let name = ''
  let value = index
  while (value > 0) {
    const remaider = (value - 1) % 26
    name = `${String.fromCharCode(65 + remaider)}${name}`
    value = Math.floor((value - remaider - 1) / 26)
  }
  return name
}

function addOverride(contentTypes: string, partName: string, contentType: string): string {
  if (contentTypes.includes(`PartName="${partName}"`)) return contentTypes
  return contentTypes.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`)
}
