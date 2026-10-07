import assert from 'node:assert/strict'
import { test } from 'node:test'
import ExcelJS from 'exceljs'
import { strFromU8, unzipSync, zipSync } from 'fflate'
import {
  annotateWorkbookXml,
  emptyAnnotations,
  type CommentSpec,
  type SparklineGroupSpec,
} from '../src/xml-postprocess.ts'

/**
 * Every part this module writes is hand-built XML, so the tests check shape
 * rather than only presence: a well-formedness scan catches an element closed in
 * the wrong order or nested inside the wrong parent, which is exactly how these
 * injections go wrong.
 */
function assertWellFormed(xml: string, label: string): void {
  const stack: string[] = []
  const tag = /<(\/?)([A-Za-z_][\w:.-]*)((?:"[^"]*"|[^>"])*?)(\/?)>/g
  for (const match of xml.matchAll(tag)) {
    const [, closing, name, , selfClosing] = match
    if (closing === '/') {
      const open = stack.pop()
      assert.equal(open, name, `${label}: </${name}> closes <${open ?? 'nothing'}>`)
    } else if (selfClosing !== '/') {
      stack.push(name!)
    }
  }
  assert.deepEqual(stack, [], `${label}: unclosed elements`)
}

async function sheetOfWorkbook(sheetName = '订单'): Promise<{ data: Uint8Array; sheetFile: string }> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet(sheetName)
  sheet.getCell('A1').value = 'x'
  sheet.getCell('B2').value = 1
  const data = new Uint8Array(await workbook.xlsx.writeBuffer())
  return { data, sheetFile: 'xl/worksheets/sheet1.xml' }
}

function readEntry(data: Uint8Array, entry: string): string | null {
  const files = unzipSync(data)
  const bytes = files[entry]
  return bytes ? strFromU8(bytes) : null
}

function comment(overrides: Partial<CommentSpec> = {}): CommentSpec {
  return { ref: 'B2', text: '看这里', author: '作者', width: 108, height: 60, ...overrides }
}

function sparkline(overrides: Partial<SparklineGroupSpec> = {}): SparklineGroupSpec {
  return {
    dataRange: '订单!B2:F3',
    locationRange: '订单!G2:G3',
    type: 'line',
    color: 'FF376092',
    negativeColor: 'FFD00000',
    markers: true,
    highColor: 'FF00B050',
    lowColor: 'FFFF0000',
    ...overrides,
  }
}

test('nothing to annotate returns the original bytes untouched', async () => {
  const { data } = await sheetOfWorkbook()
  const same = annotateWorkbookXml(data, emptyAnnotations(), new Map())
  assert.deepEqual([...same], [...data])
})

test('a comment produces every part Excel needs, and the result still opens', async () => {
  const { data, sheetFile } = await sheetOfWorkbook()
  const annotations = emptyAnnotations()
  annotations.comments.set('订单', [comment()])

  const annotated = annotateWorkbookXml(data, annotations, new Map([['订单', sheetFile]]))

  const comments = readEntry(annotated, 'xl/comments1.xml')
  const vml = readEntry(annotated, 'xl/drawings/vmlDrawing1.vml')
  const rels = readEntry(annotated, 'xl/worksheets/_rels/sheet1.xml.rels')
  const types = readEntry(annotated, '[Content_Types].xml')
  const sheet = readEntry(annotated, sheetFile)
  assert.ok(comments && vml && rels && types && sheet, 'every part must be present')

  assertWellFormed(comments, 'comments1.xml')
  assertWellFormed(vml, 'vmlDrawing1.vml')
  assertWellFormed(rels, 'sheet rels')
  assertWellFormed(types, '[Content_Types].xml')
  assertWellFormed(sheet, 'sheet1.xml')

  assert.match(sheet, /<legacyDrawing r:id="rId\d+"\/>/)
  assert.match(types, /PartName="\/xl\/comments1\.xml"/)
  assert.match(types, /Extension="vml"/)
  assert.match(rels, /relationships\/comments"/)
  assert.match(rels, /relationships\/vmlDrawing"/)

  // The real check: exceljs must be able to read what we wrote.
  const reopened = new ExcelJS.Workbook()
  await reopened.xlsx.load(Buffer.from(annotated))
  assert.equal(reopened.getWorksheet('订单')?.getCell('B2').address, 'B2')
})

test('comment text and author survive XML metacharacters', async () => {
  const { data, sheetFile } = await sheetOfWorkbook()
  const annotations = emptyAnnotations()
  annotations.comments.set('订单', [comment({ text: 'a<b>&"c\'d', author: 'A&B' })])

  const comments = readEntry(annotateWorkbookXml(data, annotations, new Map([['订单', sheetFile]])), 'xl/comments1.xml')!
  assert.match(comments, /<author>A&amp;B<\/author>/)
  assert.match(comments, /a&lt;b&gt;&amp;&quot;c'd/)
  assertWellFormed(comments, 'comments1.xml')
})

test('a comment carries the zero-based row and column Excel expects', async () => {
  const { data, sheetFile } = await sheetOfWorkbook()
  const annotations = emptyAnnotations()
  // B2 -> row 2, column B, which the VML records zero-based as 1 and 1.
  annotations.comments.set('订单', [comment({ ref: 'B2' })])

  const vml = readEntry(annotateWorkbookXml(data, annotations, new Map([['订单', sheetFile]])), 'xl/drawings/vmlDrawing1.vml')!
  assert.match(vml, /<x:Row>1<\/x:Row>/)
  assert.match(vml, /<x:Column>1<\/x:Column>/)
})

test('showFormulas lands in the sheet view whether or not one exists yet', async () => {
  const { data, sheetFile } = await sheetOfWorkbook()
  const annotations = emptyAnnotations()
  annotations.showFormulas.add('订单')

  const sheet = readEntry(annotateWorkbookXml(data, annotations, new Map([['订单', sheetFile]])), sheetFile)!
  assertWellFormed(sheet, 'sheet1.xml')
  assert.match(sheet, /<sheetView\b[^>]*showFormulas="1"/)
})

test('a sparkline group emits one sparkline per row with the matching location', async () => {
  const { data, sheetFile } = await sheetOfWorkbook()
  const annotations = emptyAnnotations()
  annotations.sparklines.set('订单', [sparkline()])

  const sheet = readEntry(annotateWorkbookXml(data, annotations, new Map([['订单', sheetFile]])), sheetFile)!
  assertWellFormed(sheet, 'sheet1.xml')

  const sparklines = [...sheet.matchAll(/<x14:sparkline>(.*?)<\/x14:sparkline>/g)].map((match) => match[1]!)
  assert.equal(sparklines.length, 2, 'one sparkline per data row')
  assert.match(sparklines[0]!, /<xm:f>订单!B2:F2<\/xm:f>/)
  assert.match(sparklines[0]!, /<xm:sqref>G2<\/xm:sqref>/)
  assert.match(sparklines[1]!, /<xm:f>订单!B3:F3<\/xm:f>/)
  assert.match(sparklines[1]!, /<xm:sqref>G3<\/xm:sqref>/)
})

test('the sparkline extension is a direct child of the worksheet extension list', async () => {
  // A worksheet may carry an extension list of its own, and extensions nest. The
  // sparkline group has to be appended to the worksheet's own list, not dropped
  // into whichever closing tag happens to come first.
  const { data, sheetFile } = await sheetOfWorkbook()
  const annotations = emptyAnnotations()
  annotations.sparklines.set('订单', [sparkline()])

  const sheet = readEntry(annotateWorkbookXml(data, annotations, new Map([['订单', sheetFile]])), sheetFile)!
  assert.match(sheet, /<extLst><ext [^>]*uri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}">/)
  // Exactly one worksheet-level list, closed after the extension.
  assert.equal(sheet.split('<extLst>').length - 1, 1)
})

test('a sparkline group merges into a worksheet that already has an extension list', async () => {
  // A worksheet can already carry extensions — data-bar and icon-set formats add
  // them. The sparkline group has to join that list rather than replace it or
  // land in the wrong one.
  const { data, sheetFile } = await sheetOfWorkbook()
  const files = unzipSync(data)
  const existing = strFromU8(files[sheetFile]!)
  const withExtLst = existing.replace(
    '</worksheet>',
    '<extLst><ext uri="{EXISTING}"><x14:thing xmlns:x14="http://x"/></ext></extLst></worksheet>',
  )
  files[sheetFile] = new TextEncoder().encode(withExtLst)
  const crafted = zipSync(files)

  const annotations = emptyAnnotations()
  annotations.sparklines.set('订单', [sparkline()])
  const sheet = readEntry(
    annotateWorkbookXml(crafted, annotations, new Map([['订单', sheetFile]])),
    sheetFile,
  )!

  assertWellFormed(sheet, 'sheet1.xml')
  assert.match(sheet, /uri="\{EXISTING\}"/, 'the existing extension must survive')
  assert.match(sheet, /uri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}"/, 'the sparkline extension must be added')
})

test('a sparkline group does not land inside a nested extension list', async () => {
  // `<extLst>` can nest: a conditional-formatting rule carries its own. Appending
  // to "the first closing extLst tag" would put the sparkline group inside that
  // rule and produce a worksheet Excel refuses to open.
  const { data, sheetFile } = await sheetOfWorkbook()
  const files = unzipSync(data)
  const existing = strFromU8(files[sheetFile]!)
  const withNested = existing.replace(
    '</worksheet>',
    '<conditionalFormatting sqref="A1:A2"><cfRule type="dataBar" priority="1">'
      + '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>'
      + '<extLst><ext uri="{NESTED}"><x14:inner xmlns:x14="http://x"/></ext></extLst>'
      + '</cfRule></conditionalFormatting></worksheet>',
  )
  files[sheetFile] = new TextEncoder().encode(withNested)
  const crafted = zipSync(files)

  const annotations = emptyAnnotations()
  annotations.sparklines.set('订单', [sparkline()])
  const sheet = readEntry(
    annotateWorkbookXml(crafted, annotations, new Map([['订单', sheetFile]])),
    sheetFile,
  )!

  assertWellFormed(sheet, 'sheet1.xml')
  assert.match(sheet, /\{NESTED\}/)
  assert.match(sheet, /\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}/)
  // The sparkline extension must sit after the conditional formatting, not inside it.
  const sparklineAt = sheet.indexOf('{05C60535-1F16-4fd2-B633-F4F36F0B64E0}')
  const cfEnd = sheet.indexOf('</conditionalFormatting>')
  assert.ok(
    sparklineAt > cfEnd,
    'the sparkline extension must be a worksheet-level sibling, not nested in a cfRule',
  )
})

test('a comment does not land inside a nested extension list either', async () => {
  // legacyDrawing must precede tableParts and extLst. Matching the first
  // `<extLst` would put it inside the conditional-formatting rule's own list.
  const { data, sheetFile } = await sheetOfWorkbook()
  const files = unzipSync(data)
  const existing = strFromU8(files[sheetFile]!)
  const withNested = existing.replace(
    '</worksheet>',
    '<conditionalFormatting sqref="A1:A2"><cfRule type="dataBar" priority="1">'
      + '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>'
      + '<extLst><ext uri="{NESTED}"><x14:inner xmlns:x14="http://x"/></ext></extLst>'
      + '</cfRule></conditionalFormatting></worksheet>',
  )
  files[sheetFile] = new TextEncoder().encode(withNested)
  const crafted = zipSync(files)

  const annotations = emptyAnnotations()
  annotations.comments.set('订单', [comment()])
  const sheet = readEntry(
    annotateWorkbookXml(crafted, annotations, new Map([['订单', sheetFile]])),
    sheetFile,
  )!

  assertWellFormed(sheet, 'sheet1.xml')
  const legacyAt = sheet.indexOf('<legacyDrawing')
  const cfEnd = sheet.indexOf('</conditionalFormatting>')
  assert.ok(legacyAt > cfEnd, 'legacyDrawing must be a worksheet-level sibling, not inside a cfRule')
})

test('a sparkline group whose rows do not line up is rejected', async () => {
  const { data, sheetFile } = await sheetOfWorkbook()
  const annotations = emptyAnnotations()
  annotations.sparklines.set('订单', [sparkline({ dataRange: '订单!B2:F4', locationRange: '订单!G2:G3' })])

  assert.throws(
    () => annotateWorkbookXml(data, annotations, new Map([['订单', sheetFile]])),
    /must match locationRange rows/,
  )
})

test('a range on another sheet is rejected', async () => {
  const { data, sheetFile } = await sheetOfWorkbook()
  const annotations = emptyAnnotations()
  annotations.sparklines.set('订单', [sparkline({ dataRange: '数据!B2:F3' })])

  assert.throws(
    () => annotateWorkbookXml(data, annotations, new Map([['订单', sheetFile]])),
    /same sheet/,
  )
})

test('annotating a sheet that is not in the workbook fails loudly', async () => {
  const { data, sheetFile } = await sheetOfWorkbook()
  for (const [label, prepare] of [
    ['comments', (a: ReturnType<typeof emptyAnnotations>) => a.comments.set('不存在', [comment()])],
    ['sparklines', (a: ReturnType<typeof emptyAnnotations>) => a.sparklines.set('不存在', [sparkline()])],
    ['showFormulas', (a: ReturnType<typeof emptyAnnotations>) => a.showFormulas.add('不存在')],
  ] as const) {
    const annotations = emptyAnnotations()
    prepare(annotations)
    assert.throws(
      () => annotateWorkbookXml(data, annotations, new Map([['订单', sheetFile]])),
      /sheet not found/,
      `${label} must fail rather than write nothing`,
    )
  }
})

test('several commented sheets each get their own parts', async () => {
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('一').getCell('A1').value = 1
  workbook.addWorksheet('二').getCell('A1').value = 2
  const data = new Uint8Array(await workbook.xlsx.writeBuffer())

  const annotations = emptyAnnotations()
  annotations.comments.set('一', [comment({ ref: 'A1' })])
  annotations.comments.set('二', [comment({ ref: 'A1', text: '第二个' })])

  const annotated = annotateWorkbookXml(data, annotations, new Map([
    ['一', 'xl/worksheets/sheet1.xml'],
    ['二', 'xl/worksheets/sheet2.xml'],
  ]))

  assert.ok(readEntry(annotated, 'xl/comments1.xml'))
  assert.ok(readEntry(annotated, 'xl/comments2.xml'))
  assert.ok(readEntry(annotated, 'xl/drawings/vmlDrawing1.vml'))
  assert.ok(readEntry(annotated, 'xl/drawings/vmlDrawing2.vml'))

  const reopened = new ExcelJS.Workbook()
  await reopened.xlsx.load(Buffer.from(annotated))
  assert.deepEqual(reopened.worksheets.map((sheet) => sheet.name), ['一', '二'])
})
