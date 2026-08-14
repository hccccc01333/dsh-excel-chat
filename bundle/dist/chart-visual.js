import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { mkdir, readdir, writeFile } from 'node:fs/promises';
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
export async function exportChartsWithExcel(path, outDir, signal) {
    await mkdir(outDir, { recursive: true });
    const scriptPath = join(tmpdir(), `vera-export-${randomUUID()}.ps1`);
    await writeFile(scriptPath, EXCEL_EXPORT_SCRIPT, 'utf8');
    await runPowerShell([
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        scriptPath,
        '-WorkbookPath',
        path,
        '-OutDir',
        outDir,
    ], signal);
    const files = (await readdir(outDir)).filter((name) => name.startsWith('chart-') && name.endsWith('.png')).sort();
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
    const scriptPath = join(tmpdir(), `vera-chart-create-${randomUUID()}.ps1`);
    await writeFile(scriptPath, CHART_CREATE_SCRIPT, 'utf8');
    await runPowerShell([
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
    const scriptPath = join(tmpdir(), `vera-chart-modify-${randomUUID()}.ps1`);
    await writeFile(scriptPath, CHART_MODIFY_SCRIPT, 'utf8');
    await runPowerShell([
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
