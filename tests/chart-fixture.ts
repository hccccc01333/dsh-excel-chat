import ExcelJS from 'exceljs'
import { strToU8, unzipSync, zipSync } from 'fflate'

export function chartXml(type: string, seriesXml: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:layout/><c:${type}>${seriesXml}</c:${type}><c:catAx/><c:valAx/></c:plotArea></c:chart>
</c:chartSpace>`
}

export const goodSeries = `<c:ser>
  <c:tx><c:strRef><c:f>Sales!$B$1</c:f></c:strRef></c:tx>
  <c:cat><c:strRef><c:f>Sales!$A$2:$A$4</c:f></c:strRef></c:cat>
  <c:val><c:numRef><c:f>Sales!$B$2:$B$4</c:f></c:numRef></c:val>
</c:ser>`

export const badSeries = `<c:ser>
  <c:cat><c:strRef><c:f>Sales!$A$2:$C$4</c:f></c:strRef></c:cat>
  <c:val><c:numRef><c:f>Sales!$X$9:$X$10</c:f></c:numRef></c:val>
</c:ser>
<c:ser>
  <c:val><c:numRef><c:f>Sales!$B$2:$B$4</c:f></c:numRef></c:val>
</c:ser>`

export const unsortedSeries = `<c:ser>
  <c:cat><c:strRef><c:f>Sales!$D$2:$D$4</c:f></c:strRef></c:cat>
  <c:val><c:numRef><c:f>Sales!$B$2:$B$4</c:f></c:numRef></c:val>
</c:ser>`

export async function buildChartWorkbook(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sales')
  sheet.getCell('A1').value = 'Date'
  sheet.getCell('A2').value = new Date('2026-01-01T00:00:00Z')
  sheet.getCell('A3').value = new Date('2026-02-01T00:00:00Z')
  sheet.getCell('A4').value = new Date('2026-03-01T00:00:00Z')
  sheet.getCell('D2').value = new Date('2026-03-01T00:00:00Z')
  sheet.getCell('D3').value = new Date('2026-02-01T00:00:00Z')
  sheet.getCell('D4').value = new Date('2026-01-01T00:00:00Z')
  sheet.getCell('B1').value = 'Revenue'
  sheet.getCell('B2').value = 100
  sheet.getCell('B3').value = 200
  sheet.getCell('B4').value = 300
  sheet.getCell('C2').value = 60
  sheet.getCell('C3').value = 120
  sheet.getCell('C4').value = 150
  const files = unzipSync(new Uint8Array(await workbook.xlsx.writeBuffer()))
  files['xl/charts/chart1.xml'] = strToU8(chartXml('lineChart', goodSeries))
  files['xl/charts/chart2.xml'] = strToU8(chartXml('barChart', badSeries))
  files['xl/charts/chart3.xml'] = strToU8(chartXml('lineChart', unsortedSeries))
  files['xl/drawings/drawing1.xml'] = strToU8(drawingXml())
  files['xl/drawings/_rels/drawing1.xml.rels'] = strToU8(drawingRels())
  files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(sheetRels())
  const contentTypes = new TextDecoder().decode(files['[Content_Types].xml'])
  const overrides = [
    '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    '<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    '<Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
  ].join('')
  files['[Content_Types].xml'] = strToU8(contentTypes.replace('</Types>', `${overrides}</Types>`))
  return Buffer.from(zipSync(files))
}

function drawingXml(): string {
  const frame = (id: number, name: string, chartId: string, col: number): string => `
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>${col + 8}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${chartId}"/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  ${frame(2, 'Chart 1', 'rIdChart1', 0)}
  ${frame(3, 'Chart 2', 'rIdChart2', 9)}
  ${frame(4, 'Chart 3', 'rIdChart3', 18)}
</xdr:wsDr>`
}

function drawingRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
  <Relationship Id="rIdChart3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
</Relationships>`
}

function sheetRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
  <Relationship Id="rIdChart3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`
}
