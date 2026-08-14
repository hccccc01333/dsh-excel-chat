import ExcelJS from 'exceljs'
import { randomUUID } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { columnToNumber, numberToColumn } from './formula.ts'
import { runPowerShell } from './chart-visual.ts'

export interface PivotValueSpec {
  column: string
  function: 'sum' | 'count' | 'average' | 'max' | 'min'
}

export interface PivotOptions {
  sheet: string
  /** Source data range including the header row, e.g. "订单!A1:F7". */
  range: string
  rows: string[]
  columns?: string[]
  filters?: string[]
  values: PivotValueSpec[]
  outputSheet?: string
}

export interface PivotResult {
  pivotSheet: string
  groups: number
  recordCount: number
}

const COM_FUNCTIONS: Record<PivotValueSpec['function'], number> = {
  sum: -4157, // xlSum
  count: -4112, // xlCount
  average: -4106, // xlAverage
  max: -4136, // xlMax
  min: -4139, // xlMin
}

const PIVOT_SCRIPT = `
param([string]$configPath)
$ErrorActionPreference = 'Stop'
$config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($config.workbookPath)
  $source = $wb.Worksheets.Item($config.sheet)
  $range = $source.Range($config.range)
  $pivotSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $source)
  $pivotSheet.Name = $config.pivotSheet
  $cache = $wb.PivotCaches().Create(1, $range)
  $pt = $pivotSheet.PivotTables().Add($cache, $pivotSheet.Range("A3"), "PivotTable1")
  for ($i = 0; $i -lt $config.rowFields.Count; $i++) {
    $pf = $pt.PivotFields($config.rowFields[$i])
    $pf.Orientation = 1
    $pf.Position = $i + 1
  }
  for ($i = 0; $i -lt $config.columnFields.Count; $i++) {
    $pf = $pt.PivotFields($config.columnFields[$i])
    $pf.Orientation = 2
    $pf.Position = $i + 1
  }
  for ($i = 0; $i -lt $config.filterFields.Count; $i++) {
    $pf = $pt.PivotFields($config.filterFields[$i])
    $pf.Orientation = 3
    $pf.Position = $i + 1
  }
  for ($i = 0; $i -lt $config.valueFields.Count; $i++) {
    $pf = $pt.PivotFields($config.valueFields[$i].name)
    $pf.Orientation = 4
    $pf.Function = $config.valueFields[$i].function
  }
  $wb.SaveAs($config.outPath, 51)
  $wb.Close($false)
  Write-Output "OK"
} finally {
  $excel.Quit()
}
`

/**
 * Create a native Excel pivot table by driving Excel COM (Windows): the cache
 * and pivot table are produced by Excel itself, so the output is always valid.
 */
export async function createPivotTable(
  inputPath: string,
  options: PivotOptions,
  outPath: string,
): Promise<PivotResult> {
  if (options.rows.length < 1) throw new Error('pivot requires at least one row field')
  if (options.values.length < 1) throw new Error('pivot requires at least one value field')
  if (process.platform !== 'win32') {
    throw new Error('native pivot tables require Windows with Microsoft Excel installed')
  }

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await readFile(inputPath) as any)
  const source = workbook.getWorksheet(options.sheet)
  if (!source) throw new Error(`sheet not found: ${options.sheet}`)
  const bang = options.range.lastIndexOf('!')
  const rangeBody = bang >= 0 ? options.range.slice(bang + 1) : options.range
  const rangeMatch = /^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/.exec(rangeBody)
  if (!rangeMatch) throw new Error(`invalid pivot source range: ${options.range}`)
  const startCol = columnToNumber(rangeMatch[1]!)
  const startRow = Number(rangeMatch[2]!)
  const endCol = columnToNumber(rangeMatch[3]!)
  const endRow = Number(rangeMatch[4]!)

  const resolveField = (letter: string, label: string): string => {
    const col = columnToNumber(letter)
    if (col < startCol || col > endCol) throw new Error(`${label} outside range: ${letter}`)
    return String(source.getCell(`${numberToColumn(col)}${startRow}`).value ?? letter)
  }
  const rowFields = options.rows.map((letter) => resolveField(letter, 'row field'))
  const columnFields = (options.columns ?? []).map((letter) => resolveField(letter, 'column field'))
  const filterFields = (options.filters ?? []).map((letter) => resolveField(letter, 'filter field'))
  const valueFields = options.values.map((value) => {
    const col = columnToNumber(value.column)
    if (col < startCol || col > endCol) throw new Error(`value column outside range: ${value.column}`)
    return {
      name: String(source.getCell(`${numberToColumn(col)}${startRow}`).value ?? value.column),
      function: COM_FUNCTIONS[value.function],
    }
  })

  const groupValues = new Set<string>()
  for (let row = startRow + 1; row <= endRow; row++) {
    const key = rowFields.map((_name, index) => {
      const letter = options.rows[index]!
      const raw = source.getCell(`${letter}${row}`).value
      return raw === null || raw === undefined ? '' : String(raw)
    }).join('|')
    groupValues.add(key)
  }

  const config = {
    workbookPath: inputPath,
    outPath,
    sheet: options.sheet,
    range: rangeBody,
    pivotSheet: options.outputSheet ?? `${options.sheet}透视`,
    rowFields,
    columnFields,
    filterFields,
    valueFields,
  }
  const configPath = join(tmpdir(), `vera-pivot-${randomUUID()}.json`)
  await writeFile(configPath, '\uFEFF' + JSON.stringify(config), 'utf8')
  const scriptPath = join(tmpdir(), `vera-pivot-${randomUUID()}.ps1`)
  await writeFile(scriptPath, '\uFEFF' + PIVOT_SCRIPT, 'utf8')
  await runPowerShell([
    '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
    '-configPath', configPath,
  ])
  return {
    pivotSheet: config.pivotSheet,
    groups: groupValues.size,
    recordCount: endRow - startRow,
  }
}
