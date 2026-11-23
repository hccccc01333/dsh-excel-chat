import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { mkdir, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
const POWERSHELL_PATH = process.env.SystemRoot
    ? join(process.env.SystemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
    : 'powershell.exe';
const EXCEL_EXPORT_SCRIPT = `
param([string]$WorkbookPath, [string]$OutDir)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($WorkbookPath, $null, $true)
  $count = 0
  try {
    foreach ($ws in $wb.Worksheets) {
      foreach ($chart in $ws.ChartObjects()) {
        $count++
        $target = Join-Path $OutDir ("chart-{0}.png" -f $count)
        $chart.Chart.Export($target, "PNG")
      }
    }
  } finally {
    $wb.Close($false)
  }
} finally {
  $excel.Quit()
}
Write-Output $count
`;
/**
 * Write the PowerShell body to a unique temp file and always remove it. The
 * scripts used to accumulate in tmpdir — every chart/PDF call left one behind.
 */
async function withPowerShellScript(script, prefix, args, signal) {
    const scriptPath = join(tmpdir(), `vera-${prefix}-${randomUUID()}.ps1`);
    await writeFile(scriptPath, script, 'utf8');
    try {
        return await runPowerShell(args(scriptPath), signal);
    }
    finally {
        await rm(scriptPath, { force: true }).catch(() => { });
    }
}
const chartIndex = (name) => Number(/chart-(\d+)\.png$/.exec(name)?.[1] ?? 0);
/**
 * Only the PNGs this export actually produced are results; older files in the
 * same outDir must not be reported as charts of the current workbook.
 */
export function selectExportedCharts(files, stdout) {
    const charts = files.filter((name) => /^chart-\d+\.png$/.test(name))
        .sort((left, right) => chartIndex(left) - chartIndex(right));
    const tail = stdout.trim().split(/\r?\n/).pop()?.trim() ?? '';
    // Number('') is 0; an empty stdout must mean "no count reported", not "zero charts".
    const count = tail === '' ? Number.NaN : Number(tail);
    return Number.isInteger(count) && count >= 0 ? charts.slice(0, count) : charts;
}
export async function exportChartsWithExcel(path, outDir, signal) {
    await mkdir(outDir, { recursive: true });
    const stdout = await withPowerShellScript(EXCEL_EXPORT_SCRIPT, 'export', (scriptPath) => [
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
        '-WorkbookPath', path,
        '-OutDir', outDir,
    ], signal);
    // The script prints how many charts it exported; use it to ignore stale PNGs
    // from an earlier export into the same folder.
    const files = selectExportedCharts(await readdir(outDir), stdout);
    return files.map((name) => join(outDir, name));
}
const CHART_TYPE_CODES = {
    column: 51, // xlColumnClustered
    line: 4, // xlLine
    pie: 5, // xlPie
    bar: 57, // xlBarClustered
    area: 1, // xlArea
};
const CHART_CREATE_SCRIPT = `
param([string]$WorkbookPath, [string]$OutPath, [string]$SheetName, [string]$Range, [int]$ChartType, [string]$ChartName, [string]$Title)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($WorkbookPath)
  $ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.Worksheets.Item(1) }
  $chart = $ws.ChartObjects().Add(0, 0, 480, 300)
  $chart.Name = $ChartName
  $chart.Chart.ChartType = $ChartType
  $chart.Chart.SetSourceData($ws.Range($Range))
  if ($Title) {
    $chart.Chart.HasTitle = $true
    $chart.Chart.ChartTitle.Text = $Title
  }
  $wb.SaveAs($OutPath, 51)
  $wb.Close($false)
} finally {
  $excel.Quit()
}
`;
const CHART_MODIFY_SCRIPT = `
param([string]$WorkbookPath, [string]$OutPath, [string]$ChartName, [int]$ChartType, [string]$Title, [string]$AxisX, [string]$AxisY, [string]$Legend)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($WorkbookPath)
  $chart = $null
  foreach ($ws in $wb.Worksheets) {
    foreach ($item in $ws.ChartObjects()) {
      if ($item.Name -eq $ChartName) { $chart = $item; break }
    }
    if ($chart) { break }
  }
  if (-not $chart) { throw "chart not found: $ChartName" }
  if ($ChartType -gt 0) { $chart.Chart.ChartType = $ChartType }
  if ($Title) {
    $chart.Chart.HasTitle = $true
    $chart.Chart.ChartTitle.Text = $Title
  }
  if ($AxisX) {
    $chart.Chart.Axes(1).HasTitle = $true
    $chart.Chart.Axes(1).AxisTitle.Text = $AxisX
  }
  if ($AxisY) {
    $chart.Chart.Axes(2).HasTitle = $true
    $chart.Chart.Axes(2).AxisTitle.Text = $AxisY
  }
  if ($Legend -ne '') {
    $chart.Chart.HasLegend = ($Legend -eq 'True')
  }
  $wb.SaveAs($OutPath, 51)
  $wb.Close($false)
} finally {
  $excel.Quit()
}
`;
/** Create a chart in an .xlsx copy using local Excel (Windows only). */
export async function createChartWithExcel(inputPath, options, outPath, signal) {
    await withPowerShellScript(CHART_CREATE_SCRIPT, 'chart-create', (scriptPath) => [
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
        '-WorkbookPath', inputPath,
        '-OutPath', outPath,
        '-SheetName', options.sheet ?? '',
        '-Range', options.range,
        '-ChartType', String(CHART_TYPE_CODES[options.type ?? 'column']),
        '-ChartName', options.name ?? 'Chart 1',
        '-Title', options.title ?? '',
    ], signal);
}
/** Modify chart parameters (type, title, legend, axis titles) in an .xlsx copy. */
export async function modifyChartWithExcel(inputPath, chartName, changes, outPath, signal) {
    await withPowerShellScript(CHART_MODIFY_SCRIPT, 'chart-modify', (scriptPath) => [
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
        '-WorkbookPath', inputPath,
        '-OutPath', outPath,
        '-ChartName', chartName,
        '-ChartType', changes.type ? String(CHART_TYPE_CODES[changes.type]) : '0',
        '-Title', changes.title ?? '',
        '-AxisX', changes.axisTitleX ?? '',
        '-AxisY', changes.axisTitleY ?? '',
        '-Legend', changes.hasLegend === undefined ? '' : String(changes.hasLegend),
    ], signal);
}
export function runPowerShell(args, signal) {
    return new Promise((resolve, reject) => {
        const child = spawn(POWERSHELL_PATH, args, { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
        let stdout = '';
        let stderr = '';
        child.stdout.on('data', (data) => { stdout += String(data); });
        child.stderr.on('data', (data) => { stderr += String(data); });
        const onAbort = () => { child.kill(); };
        signal?.addEventListener('abort', onAbort, { once: true });
        child.on('error', reject);
        child.on('close', (code) => {
            signal?.removeEventListener('abort', onAbort);
            if (code === 0)
                resolve(stdout);
            else
                reject(new Error(`PowerShell failed (${code}): ${stderr.trim()}`));
        });
    });
}
const PDF_EXPORT_SCRIPT = `
param([string]$WorkbookPath, [string]$OutPath, [string]$SheetName)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($WorkbookPath, $null, $true)
  try {
    if ($SheetName) {
      $ws = $wb.Worksheets.Item($SheetName)
      $ws.ExportAsFixedFormat(0, $OutPath)
    } else {
      $wb.ExportAsFixedFormat(0, $OutPath)
    }
  } finally {
    $wb.Close($false)
  }
} finally {
  $excel.Quit()
}
Write-Output "ok"
`;
/**
 * Export a workbook (or a single sheet) to PDF using local Excel COM.
 * Windows only; the file is opened read-only and left untouched.
 */
export async function exportWorkbookToPdf(inputPath, outPath, sheet, signal) {
    await withPowerShellScript(PDF_EXPORT_SCRIPT, 'pdf', (scriptPath) => [
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
        '-WorkbookPath', inputPath,
        '-OutPath', outPath,
        '-SheetName', sheet ?? '',
    ], signal);
}
export function buildVisionPrompt() {
    return [
        'You are the visual critic for Excel charts exported as PNG images.',
        'Check every item and return ONLY JSON: {"issues":[{"kind":"...","severity":"info|warning|critical","description":"..."}]}.',
        'Checklist:',
        '- title truncated or missing',
        '- legend overlapping or clipped',
        '- data labels overlapping',
        '- axis labels unreadable or unreasonable',
        '- chart too crowded or whitespace excessive',
        '- trends cannot be read from the chart',
    ].join('\n');
}
export function parseVisionReply(text) {
    const match = /```(?:json)?\s*([\s\S]*?)```/.exec(text);
    const body = match ? match[1].trim() : text.trim();
    const parsed = JSON.parse(body);
    if (!Array.isArray(parsed.issues)) {
        throw new Error('vision reply must contain an issues array');
    }
    return parsed.issues;
}
export function createVisionCritic(vision) {
    return async (imagePath, signal) => {
        return { imagePath, issues: parseVisionReply(await vision(imagePath, buildVisionPrompt(), signal)) };
    };
}
export async function validateChartsVisually(path, options) {
    const images = await options.exporter(path, options.outDir, options.signal);
    const reports = [];
    for (const image of images) {
        reports.push(await options.critic(image, options.signal));
    }
    return { images, reports };
}
