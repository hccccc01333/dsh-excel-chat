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
function runPowerShell(args, signal) {
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
