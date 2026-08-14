import { test } from 'node:test'
import assert from 'node:assert/strict'
import { execFile } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import {
  buildVisionPrompt,
  createChartWithExcel,
  createVisionCritic,
  exportChartsWithExcel,
  modifyChartWithExcel,
  parseVisionReply,
  validateChartsVisually,
  type VisionText,
} from '../src/chart-visual.ts'
const excelAvailable = existsSync('C:/Program Files/Microsoft Office/Root/Office16/EXCEL.EXE')
const execFileAsync = promisify(execFile)
const POWERSHELL = join(process.env.SystemRoot ?? 'C:/Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')

test('buildVisionPrompt covers the visual checklist', () => {
  const prompt = buildVisionPrompt()
  for (const keyword of ['title truncated', 'legend', 'labels overlapping', 'axis', 'crowded', 'trends']) {
    assert.match(prompt, new RegExp(keyword))
  }
})

test('parseVisionReply handles fenced JSON and returns issues', () => {
  const issues = parseVisionReply('```json\n{"issues":[{"kind":"legend","severity":"warning","description":"overlap"}]}\n```')
  assert.deepEqual(issues, [{ kind: 'legend', severity: 'warning', description: 'overlap' }])
})

test('createVisionCritic delegates to the vision function', async () => {
  const vision: VisionText = async (imagePath) => JSON.stringify({
    issues: [{ kind: 'title', severity: 'critical', description: `missing in ${imagePath}` }],
  })
  const report = await createVisionCritic(vision)('D:/fake/chart-1.png')
  assert.equal(report.issues[0]!.kind, 'title')
})

test('validateChartsVisually combines exporter and critic', async () => {
  const images = ['D:/fake/chart-1.png', 'D:/fake/chart-2.png']
  const result = await validateChartsVisually('D:/fake/book.xlsx', {
    exporter: async () => images,
    critic: async (imagePath) => ({ imagePath, issues: [] }),
    outDir: 'D:/fake/out',
  })
  assert.equal(result.images.length, 2)
  assert.equal(result.reports.length, 2)
})

test('exportChartsWithExcel exports chart PNGs using local Excel', { skip: !excelAvailable, timeout: 120000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-charts-'))
  const book = join(dir, 'charts.xlsx')
  await createChartWorkbookWithExcel(book)
  const outDir = join(dir, 'out')
  const images = await exportChartsWithExcel(book, outDir)
  assert.equal(images.length, 3)
  for (const image of images) assert.ok(existsSync(image))
})

test('createChartWithExcel adds a chart and modifyChartWithExcel changes its parameters', { skip: !excelAvailable, timeout: 180000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-chart-edit-'))
  const book = join(dir, 'charts.xlsx')
  await createChartWorkbookWithExcel(book)
  const created = join(dir, 'created.xlsx')
  await createChartWithExcel(book, {
    range: 'Sales!A1:B4',
    type: 'line',
    title: 'Line Trend',
    name: 'Chart 4',
  }, created)
  const modified = join(dir, 'modified.xlsx')
  await modifyChartWithExcel(created, 'Chart 4', {
    type: 'column',
    title: 'Column Trend',
    hasLegend: false,
    axisTitleX: 'Date',
    axisTitleY: 'Amount',
  }, modified)
  const outDir = join(dir, 'out')
  const images = await exportChartsWithExcel(modified, outDir)
  assert.equal(images.length, 4)
})

async function createChartWorkbookWithExcel(path: string): Promise<void> {
  const script = `
param([string]$WorkbookPath)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'Sales'
  $ws.Cells.Item(1, 1) = 'Date'
  $ws.Cells.Item(2, 1) = [datetime]'2026-01-01'
  $ws.Cells.Item(3, 1) = [datetime]'2026-02-01'
  $ws.Cells.Item(4, 1) = [datetime]'2026-03-01'
  $ws.Cells.Item(1, 2) = 'Revenue'
  $ws.Cells.Item(2, 2) = 100
  $ws.Cells.Item(3, 2) = 200
  $ws.Cells.Item(4, 2) = 300
  $ws.Cells.Item(1, 3) = 'Cost'
  $ws.Cells.Item(2, 3) = 60
  $ws.Cells.Item(3, 3) = 120
  $ws.Cells.Item(4, 3) = 150

  $chart1 = $ws.ChartObjects().Add(0, 0, 400, 250)
  $chart1.Name = 'Chart 1'
  $chart1.Chart.ChartType = 4
  $chart1.Chart.SetSourceData($ws.Range('A1:B4'))
  $chart1.Chart.HasTitle = $true
  $chart1.Chart.ChartTitle.Text = 'Revenue Trend'

  $chart2 = $ws.ChartObjects().Add(450, 0, 400, 250)
  $chart2.Name = 'Chart 2'
  $chart2.Chart.ChartType = 57
  $chart2.Chart.SetSourceData($ws.Range('A1:C4'))

  $chart3 = $ws.ChartObjects().Add(900, 0, 400, 250)
  $chart3.Name = 'Chart 3'
  $chart3.Chart.ChartType = 4
  $chart3.Chart.SetSourceData($ws.Range('B1:B4'))

  $wb.SaveAs($WorkbookPath, 51)
  $wb.Close($false)
} finally {
  $excel.Quit()
}`
  const scriptPath = join(tmpdir(), `vera-make-charts-${Date.now()}.ps1`)
  await writeFile(scriptPath, script, 'utf8')
  await execFileAsync(POWERSHELL, [
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    scriptPath,
    '-WorkbookPath',
    path,
  ])
}
