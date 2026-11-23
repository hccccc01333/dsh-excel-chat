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

/** A saved autoFilter criterion ExcelJS cannot express on its own. */
export interface AutoFilterColumnSpec {
  /** Zero-based column index within the autoFilter range. */
  colId: number
  /** Exact values to select (the <filters> form). */
  values?: string[]
  /** Comparison to write as <customFilters>, e.g. { operator: 'greaterThan', val: '1000' }. */
  custom?: { operator: string; val: string }
}

export interface ChartSeries {
  name: string
  ref: string
  values: number[]
}

/** A chart ExcelJS cannot create, described entirely in worksheet coordinates. */
export interface ChartSpec {
  type: 'column' | 'bar' | 'line' | 'pie'
  title?: string
  categoryRef: string
  categories: string[]
  series: ChartSeries[]
  legend?: boolean
  xAxisTitle?: string
  yAxisTitle?: string
  /** Zero-based top-left cell for the chart. */
  anchorCol: number
  anchorRow: number
  /** Drawing size in EMU (914400 per inch). */
  width: number
  height: number
}

export type PivotAggregate = 'sum' | 'average' | 'count' | 'max' | 'min'

export interface PivotCacheField {
  name: string
  /** Numeric fields carry raw values in records; text fields carry shared-item indexes. */
  numeric: boolean
  /** Shared string items for text fields. */
  items?: string[]
  min?: number
  max?: number
}

export interface PivotMeasure {
  name: string
  /** Zero-based index into PivotSpec.fields. */
  fieldIndex: number
  aggregate: PivotAggregate
  /** Aggregated values in row order. */
  values: number[]
  /** Row-by-column values when a column field is present. */
  matrix?: number[][]
}

/** A native pivot table built from a workbook range, without Excel. */
export interface PivotSpec {
  name: string
  cacheId: number
  /** Part-number suffix, kept clear of pivot parts restored from the original file. */
  partSuffix: number
  sourceSheet: string
  /** Source range body without the sheet name, e.g. "A1:B4". */
  sourceRef: string
  /** Destination sheet and its body range, e.g. "透视!A3:B6". */
  outputSheet: string
  outputRef: string
  fields: PivotCacheField[]
  /** Zero-based index of the row field in `fields`. */
  rowFieldIndex: number
  /** Shared-item index of each output row's row-field value. */
  rowItemIndices: number[]
  /** Zero-based index of the optional column field in `fields`. */
  columnFieldIndex?: number
  /** Shared-item index of each output column's column-field value. */
  columnItemIndices?: number[]
  /** Zero-based index of the optional page (filter) field in `fields`. */
  pageFieldIndex?: number
  /** Shared-item index selected by the page field. */
  pageItemIndex?: number
  measures: PivotMeasure[]
  /** One entry per source data row, one number per field. */
  records: number[][]
}

export interface WorkbookAnnotations {
  comments: Map<string, CommentSpec[]>
  sparklines: Map<string, SparklineGroupSpec[]>
  /** Sheets whose saved view should show formulas instead of their results. */
  showFormulas: Set<string>
  /** Criteria to inject into an autoFilter that ExcelJS wrote as a bare ref. */
  filterColumns: Map<string, AutoFilterColumnSpec[]>
  /** Charts to add as chart/drawing parts. */
  charts: Map<string, ChartSpec[]>
  /** Native pivot tables to add as pivotCache/pivotTable parts. */
  pivots: PivotSpec[]
}

export function emptyAnnotations(): WorkbookAnnotations {
  return {
    comments: new Map(), sparklines: new Map(), showFormulas: new Set(),
    filterColumns: new Map(), charts: new Map(), pivots: [],
  }
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
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&#39;', "'")
    // Must be last: earlier replacements must not decode the amp in &amp;lt;.
    .replaceAll('&amp;', '&')
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
 * ExcelJS writes `<autoFilter ref="A1:D9"/>` as a leaf, so the criteria live
 * only in our model. Replace that leaf with a container carrying the filter
 * columns; Excel then shows the same arrows and applies the same filter on
 * open instead of an unfiltered range.
 */
function withAutoFilterCriteria(sheetXml: string, specs: AutoFilterColumnSpec[]): string {
  const children = specs.map((spec) => {
    if (spec.values) {
      const filters = spec.values.map((value) => `<filter val="${escapeXml(value)}"/>`).join('')
      return `<filterColumn colId="${spec.colId}"><filters>${filters}</filters></filterColumn>`
    }
    const custom = spec.custom!
    return `<filterColumn colId="${spec.colId}"><customFilters>`
      + `<customFilter operator="${escapeXml(custom.operator)}" val="${escapeXml(custom.val)}"/>`
      + '</customFilters></filterColumn>'
  }).join('')
  return sheetXml.replace(/<autoFilter\b([^>]*?)\/>/g, (match, attrs: string) => `<autoFilter${attrs}>${children}</autoFilter>`)
}

/**
 * Rewrite the saved xlsx zip: inject comments parts, VML shapes, sparkline
 * extensions, and the worksheet plumbing (legacyDrawing + rels + content
 * types) they require.
 */
export function annotateWorkbookXml(data: Uint8Array, annotations: WorkbookAnnotations, sheetFileOf: Map<string, string>): Uint8Array {
  if (
    annotations.comments.size === 0 && annotations.sparklines.size === 0 &&
    annotations.showFormulas.size === 0 && annotations.filterColumns.size === 0 &&
    annotations.charts.size === 0 && annotations.pivots.length === 0
  ) return data
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
  let chartFileIndex = 0

  const nextFreeIndex = (prefix: string, extension: string): number => {
    let index = 1
    while (files[`${prefix}${index}${extension}`]) index += 1
    return index
  }

  for (const [sheetName, comments] of annotations.comments) {
    const sheetFile = sheetFileOf.get(sheetName)
    if (!sheetFile) throw new Error(`sheet not found for comments: ${sheetName}`)
    const commentsFile = `xl/comments${nextFreeIndex('xl/comments', '.xml')}.xml`
    files[commentsFile] = strToU8(commentsXml(comments))
    contentTypes = addOverride(contentTypes, `/${commentsFile}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml')
    if (!/<Default Extension="vml"/.test(contentTypes)) {
      contentTypes = contentTypes.replace(
        /(<Types[^>]*>)/,
        '$1<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
      )
    }
    const vmlFile = `xl/drawings/vmlDrawing${nextFreeIndex('xl/drawings/vmlDrawing', '.vml')}.vml`
    const vmlXml = commentVml(comments)
    files[vmlFile] = strToU8(vmlXml)
    patchSheetForComments(files, sheetFile, commentsFile, vmlFile, vmlXml)
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

  for (const [sheetName, specs] of annotations.filterColumns) {
    const sheetFile = sheetFileOf.get(sheetName)
    if (!sheetFile) throw new Error(`sheet not found for filterColumns: ${sheetName}`)
    const xml = strFromU8(files[sheetFile] ?? new Uint8Array(0))
    const patched = withAutoFilterCriteria(xml, specs)
    if (patched === xml) throw new Error(`no autoFilter found in ${sheetFile} for criteria on ${sheetName}`)
    files[sheetFile] = strToU8(patched)
  }

  for (const [sheetName, specs] of annotations.charts) {
    const sheetFile = sheetFileOf.get(sheetName)
    if (!sheetFile) throw new Error(`sheet not found for charts: ${sheetName}`)
    const sheetRelsFile = sheetFile.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
    let sheetRels = strFromU8(files[sheetRelsFile] ?? new Uint8Array(0)) || emptyRelationships()

    // A sheet may already carry a drawing from an inserted image or a previous
    // edit. Reuse that part instead of overwriting it and orphaning its rels.
    const declaredDrawing = [...sheetRels.matchAll(/<Relationship\b[^>]*?\/>/g)]
      .map((match) => match[0])
      .find((rel) => /\/drawing$/.test(attributeOf(rel, 'Type') ?? ''))
    const reusingDrawing = Boolean(declaredDrawing)
    let drawingFile: string
    let drawingRelsFile: string
    let drawingRels: string
    if (declaredDrawing) {
      drawingFile = resolvePartPath('xl/worksheets', attributeOf(declaredDrawing, 'Target') ?? '')
      drawingRelsFile = drawingRelsFileOf(drawingFile)
      drawingRels = strFromU8(files[drawingRelsFile] ?? new Uint8Array(0)) || emptyRelationships()
    } else {
      drawingFile = `xl/drawings/drawing${nextFreeIndex('xl/drawings/drawing', '.xml')}.xml`
      drawingRelsFile = drawingRelsFileOf(drawingFile)
      drawingRels = emptyRelationships()
      files[drawingFile] = strToU8(emptyDrawingXml())
    }

    const chartRids: string[] = []
    for (const spec of specs) {
      chartFileIndex = nextFreeIndex('xl/charts/chart', '.xml')
      const chartFile = `xl/charts/chart${chartFileIndex}.xml`
      files[chartFile] = strToU8(chartXml(spec))
      const rid = nextRelationshipId(drawingRels)
      drawingRels = appendRelationship(
        drawingRels, rid, `${ns.rels}/chart`, `../charts/chart${chartFileIndex}.xml`,
      )
      chartRids.push(rid)
      contentTypes = addOverride(
        contentTypes, `/${chartFile}`,
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      )
    }

    const existingDrawingXml = strFromU8(files[drawingFile] ?? new Uint8Array(0))
    const shapeIds = [...existingDrawingXml.matchAll(/<xdr:cNvPr id="(\d+)"/g)]
      .map((match) => Number(match[1]))
    const anchors = drawingAnchors(specs, chartRids, Math.max(1, ...shapeIds) + 1)
    files[drawingFile] = strToU8(reusingDrawing
      ? existingDrawingXml.replace('</xdr:wsDr>', `${anchors}</xdr:wsDr>`)
      : wrapDrawingXml(anchors))
    files[drawingRelsFile] = strToU8(drawingRels)
    contentTypes = addOverride(
      contentTypes, `/${drawingFile}`,
      'application/vnd.openxmlformats-officedocument.drawing+xml',
    )

    if (!reusingDrawing) {
      const sheetRid = nextRelationshipId(sheetRels)
      const target = `../drawings/${drawingFile.slice(drawingFile.lastIndexOf('/') + 1)}`
      sheetRels = appendRelationship(sheetRels, sheetRid, `${ns.rels}/drawing`, target)
      files[sheetRelsFile] = strToU8(sheetRels)
      const sheetXml = strFromU8(files[sheetFile] ?? new Uint8Array(0))
      files[sheetFile] = strToU8(insertDrawingElement(sheetXml, sheetRid))
    }
  }

  if (annotations.pivots.length > 0) {
    contentTypes = addPivotParts(files, contentTypes, sheetFileOf, annotations.pivots)
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
function patchSheetForComments(
  files: Record<string, Uint8Array>,
  sheetFile: string,
  commentsFile: string,
  vmlFile: string,
  vmlXml: string,
): void {
  const sheetPath = sheetFile
  const relsPath = `${sheetPath.replace('xl/worksheets/', 'xl/worksheets/_rels/')}.rels`
  const commentsRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
  const vmlRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'

  let rels = files[relsPath] ? strFromU8(files[relsPath]) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let nextRid = 1
  for (const match of rels.matchAll(/Id="rId(\d+)"/g)) nextRid = Math.max(nextRid, Number(match[1]) + 1)
  const commentsRid = `rId${nextRid}`
  const existingVmlRel = [...rels.matchAll(/<Relationship\b[^>]*\/>/g)]
    .map((match) => match[0])
    .find((rel) => vmlRelType === (attributeOf(rel, 'Type') ?? ''))
  let vmlRid = `rId${nextRid + 1}`
  let addedRels = `<Relationship Id="${commentsRid}" Type="${commentsRelType}" Target="../${commentsFile.replace('xl/', '')}"/>`
  if (existingVmlRel) {
    vmlRid = attributeOf(existingVmlRel, 'Id') ?? vmlRid
    const existingVmlFile = resolvePartPath('xl/worksheets', attributeOf(existingVmlRel, 'Target') ?? '')
    const existingVml = strFromU8(files[existingVmlFile] ?? new Uint8Array(0))
    const ids = [...existingVml.matchAll(/_x0000_s(\d+)/g)].map((match) => Number(match[1]))
    let nextShapeId = Math.max(0, ...ids) + 1
    const shapes = [...vmlXml.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)].map((match) =>
      match[0].replace(/_x0000_s\d+/, `_x0000_s${nextShapeId++}`)).join('')
    files[existingVmlFile] = strToU8(existingVml.replace('</xml>', `${shapes}</xml>`))
  } else {
    addedRels += `<Relationship Id="${vmlRid}" Type="${vmlRelType}" Target="../${vmlFile.replace('xl/', '')}"/>`
  }
  rels = rels.replace('</Relationships>', `${addedRels}</Relationships>`)
  files[relsPath] = strToU8(rels)

  const xml = strFromU8(files[sheetPath] ?? new Uint8Array(0))
  const legacy = `<legacyDrawing r:id="${vmlRid}"/>`
  let patched = xml
  if (!/<legacyDrawing\b/.test(xml)) {
    // legacyDrawing sits after <drawing> and before tableParts, which itself
    // precedes extLst. Inserting at whichever of those two comes first keeps
    // that order. Matching the first `<extLst` instead would drop it inside a
    // nested list — the one a conditional-formatting rule carries — and make
    // the sheet invalid.
    const drawingEnd = /<drawing\b[^>]*\/>/g.exec(xml)
    const afterDrawing = drawingEnd ? drawingEnd.index + drawingEnd[0].length : -1
    const bounds = [afterDrawing, worksheetExtListStart(xml), xml.indexOf('<tableParts')].filter((at) => at >= 0)
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

function chartTitleXml(text: string): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>`
    + `<a:t>${escapeXml(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
}

function chartXml(spec: ChartSpec): string {
  const title = spec.title ? chartTitleXml(spec.title) : ''
  const categories = spec.categories.map((value, index) =>
    `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`).join('')
  const series = spec.series.map((item, index) => {
    const values = item.values.map((value, point) =>
      `<c:pt idx="${point}"><c:v>${value}</c:v></c:pt>`).join('')
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>`
      + `<c:tx><c:v>${escapeXml(item.name)}</c:v></c:tx>`
      + `<c:cat><c:strRef><c:f>${escapeXml(spec.categoryRef)}</c:f><c:strCache>`
      + `<c:ptCount val="${spec.categories.length}"/>${categories}</c:strCache></c:strRef></c:cat>`
      + `<c:val><c:numRef><c:f>${escapeXml(item.ref)}</c:f><c:numCache>`
      + `<c:formatCode>General</c:formatCode><c:ptCount val="${item.values.length}"/>${values}`
      + '</c:numCache></c:numRef></c:val></c:ser>'
  }).join('')
  const xTitle = spec.xAxisTitle ? chartTitleXml(spec.xAxisTitle) : ''
  const yTitle = spec.yAxisTitle ? chartTitleXml(spec.yAxisTitle) : ''
  const axes = spec.type === 'pie' ? '' : `
    <c:catAx><c:axId val="111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/><c:axPos val="b"/>${xTitle}<c:crossAx val="222222"/></c:catAx>
    <c:valAx><c:axId val="222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/><c:axPos val="l"/>${yTitle}<c:numFmt formatCode="General" sourceLinked="1"/>
      <c:crossAx val="111111"/></c:valAx>`
  let plot: string
  if (spec.type === 'pie') {
    plot = `<c:pieChart>${series}<c:firstSliceAng val="0"/></c:pieChart>`
  } else if (spec.type === 'line') {
    plot = `<c:lineChart><c:grouping val="standard"/>${series}<c:marker val="1"/>`
      + '<c:axId val="111111"/><c:axId val="222222"/></c:lineChart>'
  } else {
    plot = `<c:barChart><c:barDir val="${spec.type === 'bar' ? 'bar' : 'col'}"/>`
      + `<c:grouping val="clustered"/>${series}<c:gapWidth val="150"/>`
      + '<c:axId val="111111"/><c:axId val="222222"/></c:barChart>'
  }
  const legend = spec.legend ?? (spec.series.length > 1)
    ? '<c:legend><c:legendPos val="r"/></c:legend>'
    : ''
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
    + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + `<c:chart>${title}<c:plotArea><c:layout/>${plot}${axes}</c:plotArea>${legend}`
    + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>'
}

function drawingAnchors(specs: ChartSpec[], rids: string[], firstShapeId: number): string {
  return specs.map((spec, index) => `
    <xdr:oneCellAnchor>
      <xdr:from><xdr:col>${spec.anchorCol}</xdr:col><xdr:colOff>0</xdr:colOff>
        <xdr:row>${spec.anchorRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
      <xdr:ext cx="${spec.width}" cy="${spec.height}"/>
      <xdr:graphicFrame macro="">
        <xdr:nvGraphicFramePr><xdr:cNvPr id="${firstShapeId + index}" name="Chart ${firstShapeId + index}"/>
          <xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
        <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            r:id="${rids[index]}"/>
        </a:graphicData></a:graphic>
      </xdr:graphicFrame><xdr:clientData/>
    </xdr:oneCellAnchor>`).join('')
}

const drawingOpen = '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
  + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'

function emptyDrawingXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + drawingOpen + '</xdr:wsDr>'
}

function wrapDrawingXml(anchors: string): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + drawingOpen + anchors + '</xdr:wsDr>'
}

function emptyRelationships(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<Relationships xmlns="${ns.packageRel}"/>`
}

function drawingRelsFileOf(drawingFile: string): string {
  return drawingFile.replace('xl/drawings/', 'xl/drawings/_rels/') + '.rels'
}

function resolvePartPath(base: string, target: string): string {
  const parts = base.split('/')
  for (const segment of target.split('/')) {
    if (segment === '' || segment === '.') continue
    if (segment === '..') parts.pop()
    else parts.push(segment)
  }
  return parts.join('/')
}

/** CT_Worksheet puts drawing before legacyDrawing, tableParts and extLst. */
function insertDrawingElement(xml: string, rid: string): string {
  const element = `<drawing r:id="${rid}"/>`
  const points = [
    xml.indexOf('<drawing'), xml.indexOf('<legacyDrawing'), xml.indexOf('<tableParts'),
    worksheetExtListStart(xml),
  ].filter((at) => at >= 0)
  const at = points.length > 0 ? Math.min(...points) : -1
  return at >= 0 ? xml.slice(0, at) + element + xml.slice(at) : xml.replace('</worksheet>', `${element}</worksheet>`)
}

function pivotCacheDefinitionXml(spec: PivotSpec): string {
  const cacheFields = spec.fields.map((field) => {
    if (field.numeric) {
      return `<cacheField name="${escapeXml(field.name)}" numFmtId="0">`
        + '<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"'
        + ` containsInteger="1" minValue="${field.min ?? 0}" maxValue="${field.max ?? 0}"/></cacheField>`
    }
    const items = (field.items ?? []).map((value) => `<s v="${escapeXml(value)}"/>`).join('')
    return `<cacheField name="${escapeXml(field.name)}" numFmtId="0">`
      + `<sharedItems count="${field.items?.length ?? 0}">${items}</sharedItems></cacheField>`
  }).join('')
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<pivotCacheDefinition xmlns="${ns.main}" xmlns:r="${ns.rels}" r:id="rId1"`
    + ` refreshOnLoad="1" refreshedVersion="7" createdVersion="3"`
    + ` minRefreshableVersion="3" recordCount="${spec.records.length}">`
    + `<cacheSource type="worksheet"><worksheetSource ref="${escapeXml(spec.sourceRef)}" sheet="${escapeXml(spec.sourceSheet)}"/></cacheSource>`
    + `<cacheFields count="${spec.fields.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`
}

function pivotCacheRecordsXml(spec: PivotSpec): string {
  const records = spec.records.map((record) => `<r>${record.map((value, index) =>
    spec.fields[index]!.numeric ? `<n v="${value}"/>` : `<x v="${value}"/>`).join('')}</r>`).join('')
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<pivotCacheRecords xmlns="${ns.main}" count="${spec.records.length}">${records}</pivotCacheRecords>`
}

function pivotTableXml(spec: PivotSpec): string {
  const body = spec.outputRef.slice(spec.outputRef.lastIndexOf('!') + 1)
  const measureFields = new Set(spec.measures.map((measure) => measure.fieldIndex))
  const pivotFields = spec.fields.map((_field, index) => {
    if (index === spec.rowFieldIndex) {
      const items = spec.rowItemIndices.map((value) => `<item x="${value}"/>`).join('') + '<item t="default"/>'
      return `<pivotField axis="axisRow" showAll="0"><items count="${spec.rowItemIndices.length + 1}">${items}</items></pivotField>`
    }
    if (index === spec.columnFieldIndex) return '<pivotField axis="axisCol" showAll="0"/>'
    if (index === spec.pageFieldIndex) {
      const items = (_field.items ?? []).map((_value, itemIndex) => `<item x="${itemIndex}"/>`).join('')
        + '<item t="default"/>'
      return `<pivotField showAll="0"><items count="${(_field.items ?? []).length + 1}">${items}</items></pivotField>`
    }
    if (measureFields.has(index)) return '<pivotField dataField="1" showAll="0"/>'
    return '<pivotField showAll="0"/>'
  }).join('')
  const crossTab = spec.columnFieldIndex !== undefined
  const rowItems = spec.rowItemIndices.map((value) => `<i><x v="${value}"/></i>`).join('')
    + (crossTab ? '' : '<i t="grand"><x/></i>')
  // A column field occupies the column area directly; without one, several
  // measures use Excel's synthetic data field (-2) instead.
  const multi = spec.measures.length > 1
  let colFields = '<colFields count="0"/>'
  let colItems = '<colItems count="1"><i/></colItems>'
  if (crossTab) {
    const items = spec.columnItemIndices ?? []
    colFields = `<colFields count="1"><field x="${spec.columnFieldIndex}"/></colFields>`
    colItems = `<colItems count="${items.length}">${items.map((value) =>
      `<i><x v="${value}"/></i>`).join('')}</colItems>`
  } else if (multi) {
    colFields = '<colFields count="1"><field x="-2"/></colFields>'
    colItems = `<colItems count="${spec.measures.length}">${spec.measures.map((_measure, index) =>
      `<i><x v="${index}"/></i>`).join('')}</colItems>`
  }
  const grandTotals = crossTab ? ' rowGrandTotals="0" colGrandTotals="0"' : ''
  const pageFields = spec.pageFieldIndex === undefined ? ''
    : `<pageFields count="1"><pageField fld="${spec.pageFieldIndex}" item="${spec.pageItemIndex ?? 0}"/></pageFields>`
  const dataFields = spec.measures.map((measure) => {
    const subtotal = measure.aggregate === 'sum' ? '' : ` subtotal="${measure.aggregate}"`
    return `<dataField name="${escapeXml(measure.name)}" fld="${measure.fieldIndex}"`
      + `${subtotal} baseField="0" baseItem="0"/>`
  }).join('')
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<pivotTableDefinition xmlns="${ns.main}" name="${escapeXml(spec.name)}"`
    + ` cacheId="${spec.cacheId}" dataCaption="值" updatedVersion="7"`
    + ` minRefreshableVersion="3" createdVersion="3"`
    + ' useAutoFormatting="1" itemPrintTitles="1"'
    + ' applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0"'
    + ' applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1"'
    + ' indent="0" outline="1" outlineData="1" multipleFieldFilters="0"' + grandTotals + '>'
    + `<location ref="${escapeXml(body)}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>`
    + `<pivotFields count="${spec.fields.length}">${pivotFields}</pivotFields>`
    + `<rowFields count="1"><field x="${spec.rowFieldIndex}"/></rowFields>`
    + `<rowItems count="${rowItems === '' ? 0 : spec.rowItemIndices.length + (crossTab ? 0 : 1)}">${rowItems}</rowItems>`
    + colFields + colItems + pageFields
    + `<dataFields count="${spec.measures.length}">${dataFields}</dataFields>`
    + '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1"'
    + ' showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
    + '</pivotTableDefinition>'
}

const ns = {
  main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  rels: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  packageRel: 'http://schemas.openxmlformats.org/package/2006/relationships',
}

function nextRelationshipId(rels: string): string {
  const used = [...rels.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]))
  return `rId${Math.max(0, ...used) + 1}`
}

function appendRelationship(rels: string, id: string, type: string, target: string): string {
  const relationship = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`
  return rels.includes('</Relationships>')
    ? rels.replace('</Relationships>', `${relationship}</Relationships>`)
    : rels.replace(/\/>\s*$/, `>${relationship}</Relationships>`)
}

function addPivotParts(
  files: Record<string, Uint8Array>,
  contentTypes: string,
  sheetFileOf: Map<string, string>,
  specs: PivotSpec[],
): string {
  const workbookFile = 'xl/workbook.xml'
  const workbookRelsFile = 'xl/_rels/workbook.xml.rels'
  let workbookXml = strFromU8(files[workbookFile] ?? new Uint8Array(0))
  let workbookRels = strFromU8(files[workbookRelsFile] ?? new Uint8Array(0))
  const cacheRefs: string[] = []

  specs.forEach((spec, index) => {
    const suffix = spec.partSuffix || index + 1
    const cacheFile = `xl/pivotCache/pivotCacheDefinition${suffix}.xml`
    const recordsFile = `xl/pivotCache/pivotCacheRecords${suffix}.xml`
    const tableFile = `xl/pivotTables/pivotTable${suffix}.xml`
    files[cacheFile] = strToU8(pivotCacheDefinitionXml(spec))
    files[recordsFile] = strToU8(pivotCacheRecordsXml(spec))
    files[tableFile] = strToU8(pivotTableXml(spec))
    files[`xl/pivotCache/_rels/pivotCacheDefinition${suffix}.xml.rels`] = strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Relationships xmlns="${ns.packageRel}">`
      + `<Relationship Id="rId1" Type="${ns.rels}/pivotCacheRecords" Target="pivotCacheRecords${suffix}.xml"/>`
      + '</Relationships>',
    )
    files[`xl/pivotTables/_rels/pivotTable${suffix}.xml.rels`] = strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Relationships xmlns="${ns.packageRel}">`
      + `<Relationship Id="rId1" Type="${ns.rels}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${suffix}.xml"/>`
      + '</Relationships>',
    )
    for (const [part, type] of [
      [cacheFile, 'pivotCacheDefinition'], [recordsFile, 'pivotCacheRecords'], [tableFile, 'pivotTable'],
    ] as const) {
      contentTypes = addOverride(
        contentTypes, `/${part}`,
        `application/vnd.openxmlformats-officedocument.spreadsheetml.${type}+xml`,
      )
    }

    const cacheRel = nextRelationshipId(workbookRels)
    workbookRels = appendRelationship(workbookRels, cacheRel, `${ns.rels}/pivotCacheDefinition`, `pivotCache/pivotCacheDefinition${suffix}.xml`)
    cacheRefs.push(`<pivotCache cacheId="${spec.cacheId}" r:id="${cacheRel}"/>`)

    const outputFile = sheetFileOf.get(spec.outputSheet)
    if (!outputFile) throw new Error(`sheet not found for pivot output: ${spec.outputSheet}`)
    const relsFile = outputFile.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
    const sheetRels = files[relsFile]
      ? strFromU8(files[relsFile])
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${ns.packageRel}"/>`
    const tableRel = nextRelationshipId(sheetRels)
    files[relsFile] = strToU8(appendRelationship(sheetRels, tableRel, `${ns.rels}/pivotTable`, `../pivotTables/pivotTable${suffix}.xml`))
  })

  const existingCaches = /<pivotCaches\b[^>]*>([\s\S]*?)<\/pivotCaches>/.exec(workbookXml)
  if (existingCaches) {
    workbookXml = workbookXml.replace(
      existingCaches[0],
      existingCaches[0].replace('</pivotCaches>', `${cacheRefs.join('')}</pivotCaches>`),
    )
  } else {
    const caches = `<pivotCaches>${cacheRefs.join('')}</pivotCaches>`
    workbookXml = /<extLst\b/.test(workbookXml)
      ? workbookXml.replace(/<extLst\b/, `${caches}<extLst`)
      : workbookXml.replace('</workbook>', `${caches}</workbook>`)
  }
  if (!/xmlns:r\s*=/.test(workbookXml)) {
    workbookXml = workbookXml.replace(/<workbook\b/, '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
  }
  files[workbookFile] = strToU8(workbookXml)
  files[workbookRelsFile] = strToU8(workbookRels)
  return contentTypes
}

function addOverride(contentTypes: string, partName: string, contentType: string): string {
  if (contentTypes.includes(`PartName="${partName}"`)) return contentTypes
  return contentTypes.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`)
}
