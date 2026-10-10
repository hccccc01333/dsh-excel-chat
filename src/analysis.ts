import { readFile } from 'node:fs/promises'
import { t } from './i18n.ts'
import { buildWorkbookSemanticProfile, type SemanticColumn } from './semantic.ts'
import { readWorkbookCells } from './workbook.ts'

/**
 * Analytical findings: what the numbers say, not whether they are clean.
 *
 * `insight.ts` answers "is this data trustworthy" — missing values, duplicates, outliers.
 * It says nothing about what the data *means*, which is the question people actually
 * open a spreadsheet with. Measured before writing any of this: a sheet with a clean
 * 100→260 ramp and one region four times the others produced **zero** findings and no
 * mention of either pattern.
 *
 * Two findings, both deterministic and both stated with the numbers they came from, so
 * a reader can check the claim rather than trust it:
 *
 *   - **trend** — a measure rising or falling across periods of a time column
 *   - **concentration** — a few categories carrying most of a measure
 *
 * No statistics beyond sums and shares. A finding that needs a model to explain itself
 * does not belong in a report a person is expected to act on.
 */
export type AnalysisKind = 'trend' | 'concentration'

/**
 * The row the pipeline treats as the header. Skipped everywhere values are read: the
 * header's text is not data, and counting it skewed a ratio that decides whether a
 * column is a time column at all.
 */
const HEADER_ROW = 1

export interface AnalysisFinding {
  kind: AnalysisKind
  /** Where it came from, so the claim can be traced back to columns. */
  sheet: string
  message: string
  /** The numbers behind the message. Anything stated is in here. */
  evidence: Record<string, string | number>
}

/** A column of a sheet, reduced to what the analysis needs: its role and its cells. */
interface Column {
  column: string
  header: string
  role: SemanticColumn['role']
}

interface Sheet {
  sheet: string
  columns: Column[]
  /** Row number → column letter → raw text. */
  rows: Map<number, Map<string, string>>
}

/** `SHEET!A1` → `{ sheet: 'SHEET', column: 'A', row: 1 }`. */
function parseCellId(id: string): { sheet: string; column: string; row: number } | undefined {
  const match = /^(.*)!([A-Z]+)(\d+)$/.exec(id)
  if (!match) return undefined
  return { sheet: match[1]!, column: match[2]!, row: Number(match[3]) }
}

function numberFrom(text: string | undefined): number | undefined {
  if (text === undefined) return undefined
  const cleaned = text.replace(/[,\s¥$€£]/g, '')
  if (cleaned === '' || !/^-?\d*\.?\d+(e[+-]?\d+)?$/i.test(cleaned)) return undefined
  const value = Number(cleaned)
  return Number.isFinite(value) ? value : undefined
}

/**
 * Read one sheet into rows, keeping only columns the semantic pass gave a role to.
 *
 * The first row is treated as the header row, matching how the rest of the pipeline
 * detects tables.
 */
function sheetRows(
  cells: Record<string, string>,
  sheet: string,
  columns: Column[],
): Sheet {
  const rows = new Map<number, Map<string, string>>()
  const wanted = new Set(columns.map((column) => column.column))
  for (const [id, content] of Object.entries(cells)) {
    const parsed = parseCellId(id)
    if (!parsed || parsed.sheet !== sheet || !wanted.has(parsed.column)) continue
    const row = rows.get(parsed.row) ?? new Map<string, string>()
    row.set(parsed.column, content)
    rows.set(parsed.row, row)
  }
  return { sheet, columns, rows }
}

/**
 * A period label sortable as written.
 *
 * `2026-01` sorts correctly as a string; `2026年1月` does not, and neither does
 * `1月/2月/…/10月`. Both are normalised to a zero-padded form before sorting, because a
 * trend computed over a wrongly ordered series is worse than no trend.
 */
function periodKey(label: string): string {
  const yearMonth = /^(\d{4})[-/年.](\d{1,2})/.exec(label)
  if (yearMonth) return `${yearMonth[1]}-${yearMonth[2]!.padStart(2, '0')}`
  const monthOnly = /^(\d{1,2})\s*月/.exec(label)
  if (monthOnly) return `${monthOnly[1]!.padStart(2, '0')}`
  const quarter = /^(\d{4})[-/]?Q([1-4])$/i.exec(label)
  if (quarter) return `${quarter[1]}-Q${quarter[2]}`
  return label
}

/**
 * True when a label looks like a period.
 *
 * This asks whether the label *matches* a period shape — not whether normalising it
 * changed anything. The first version compared `periodKey(label)` against `label`, which
 * is false for an already-canonical `2026-01`: every well-formed period was classified
 * as not-a-period, so the trend never fired.
 */
function isPeriodLike(label: string): boolean {
  return /^(\d{4})[-/年.]\d{1,2}/.test(label) || /^\d{1,2}\s*月/.test(label) || /^(\d{4})[-/]?Q[1-4]$/i.test(label)
}

function percent(value: number): string {
  return `${Math.round(value * 1000) / 10}%`
}

/** Sum a measure by period, in period order, keeping only periods that have a value. */
function seriesByPeriod(sheet: Sheet, timeColumn: Column, measure: Column): Array<{ period: string; total: number; rows: number }> {
  // The key sorts, the label displays. They have to be separate fields: sorting by the
  // label puts `10月` before `3月`, and the first version did exactly that — it reported
  // an "upward trend" whose own numbers went from 260 down to 165.
  const totals = new Map<string, { key: string; label: string; total: number; rows: number }>()
  for (const [rowNumber, row] of sheet.rows) {
    if (rowNumber === HEADER_ROW) continue
    const label = row.get(timeColumn.column)
    const value = numberFrom(row.get(measure.column))
    if (label === undefined || label.trim() === '' || value === undefined) continue
    const key = periodKey(label.trim())
    const entry = totals.get(key) ?? { key, label: label.trim(), total: 0, rows: 0 }
    entry.total += value
    entry.rows += 1
    totals.set(key, entry)
  }
  return [...totals.values()]
    .sort((left, right) => (left.key < right.key ? -1 : left.key > right.key ? 1 : 0))
    .map((entry) => ({ period: entry.label, total: entry.total, rows: entry.rows }))
}

/**
 * Report a trend only when the series is mostly one-directional.
 *
 * A rising series with one dip is a trend; a series that wanders is not, and calling it
 * one would be the report inventing a story. The bar is two thirds of steps in the same
 * direction, and the message states how many of how many so the reader can judge.
 *
 * Not 80%: a five-period series has four steps, so 80% would demand a perfectly
 * monotonic run and reject a real trend because of one bad month. Two thirds separates
 * that (3 of 4) from a series that alternates (2 of 4).
 */
function trendFinding(sheet: Sheet, timeColumn: Column, measure: Column): AnalysisFinding | undefined {
  const series = seriesByPeriod(sheet, timeColumn, measure)
  if (series.length < 3) return undefined

  let up = 0
  let down = 0
  for (let index = 1; index < series.length; index++) {
    const delta = series[index]!.total - series[index - 1]!.total
    if (delta > 0) up++
    else if (delta < 0) down++
  }
  const steps = up + down
  if (steps < 2) return undefined
  const direction = up >= down ? 'up' : 'down'
  const consistent = direction === 'up' ? up : down
  if (consistent / steps < 2 / 3) return undefined

  const first = series[0]!
  const last = series[series.length - 1]!
  if (first.total === 0) return undefined
  const change = (last.total - first.total) / Math.abs(first.total)

  // The step majority and the end-to-end change have to agree.
  //
  // They are computed from different things — one counts steps, the other compares the
  // ends — and they can disagree: a series of three rises then a collapse is 3-of-4
  // steps "up" with a net change of -74%. Reported, that reads as "an upward trend,
  // from 1000 down to 260", which is the report contradicting itself. When they
  // disagree the series is mixed, and mixed is not a trend.
  const netDirection = change > 0 ? 'up' : change < 0 ? 'down' : undefined
  if (netDirection === undefined || netDirection !== direction) return undefined

  // The last period has to be the extreme, in the direction claimed.
  //
  // Without this, 100 → 300 → 600 → 900 → 150 reports as "an upward trend, +50%" — the
  // arithmetic is right and the reader is misled, because the series collapsed from its
  // peak. Requiring the final value to be the maximum (or minimum) rules that out while
  // still allowing a single dip, which is what a real trend looks like.
  const totals = series.map((entry) => entry.total)
  const extreme = direction === 'up' ? Math.max(...totals) : Math.min(...totals)
  if (last.total !== extreme) return undefined

  return {
    kind: 'trend',
    sheet: sheet.sheet,
    message: t('{sheet} 的 {measure} 按 {time} 呈{direction}趋势：从 {first} 的 {from} 到 {last} 的 {to}（{change}），{consistent}/{steps} 期同向。', {
      sheet: sheet.sheet,
      measure: measure.header,
      time: timeColumn.header,
      direction: direction === 'up' ? t('上升') : t('下降'),
      first: first.period,
      from: first.total,
      last: last.period,
      to: last.total,
      change: `${change >= 0 ? '+' : ''}${percent(change)}`,
      consistent,
      steps,
    }),
    evidence: {
      sheet: sheet.sheet,
      measure: measure.header,
      timeColumn: timeColumn.header,
      direction,
      periods: series.length,
      consistentSteps: consistent,
      steps,
      firstPeriod: first.period,
      firstTotal: first.total,
      lastPeriod: last.period,
      lastTotal: last.total,
      changePercent: Math.round(change * 1000) / 10,
    },
  }
}

/**
 * Report concentration only when it is actually concentrated.
 *
 * "The biggest category is 30% of the total" is not a finding — it is what a balanced
 * table looks like. The thresholds are a 50% top-1 share or an 80% top-3 share, and
 * only with at least four categories, so a two-row table cannot look concentrated.
 */
function concentrationFinding(sheet: Sheet, dimension: Column, measure: Column): AnalysisFinding | undefined {
  const totals = new Map<string, number>()
  for (const [rowNumber, row] of sheet.rows) {
    if (rowNumber === HEADER_ROW) continue
    const key = row.get(dimension.column)
    const value = numberFrom(row.get(measure.column))
    if (key === undefined || key.trim() === '' || value === undefined) continue
    totals.set(key.trim(), (totals.get(key.trim()) ?? 0) + value)
  }
  const sorted = [...totals.entries()].sort((left, right) => right[1] - left[1])
  if (sorted.length < 4) return undefined
  const total = sorted.reduce((sum, [, value]) => sum + value, 0)
  if (total <= 0) return undefined

  const top1 = sorted[0]!
  const top3 = sorted.slice(0, 3).reduce((sum, [, value]) => sum + value, 0)
  const top1Share = top1[1] / total
  const top3Share = top3 / total
  if (top1Share < 0.5 && top3Share < 0.8) return undefined

  return {
    kind: 'concentration',
    sheet: sheet.sheet,
    message: t('{sheet} 的 {measure} 集中在少数 {dimension}：{top} 占 {topShare}，前三占 {top3Share}（共 {count} 项）。', {
      sheet: sheet.sheet,
      measure: measure.header,
      dimension: dimension.header,
      top: top1[0],
      topShare: percent(top1Share),
      top3Share: percent(top3Share),
      count: sorted.length,
    }),
    evidence: {
      sheet: sheet.sheet,
      measure: measure.header,
      dimension: dimension.header,
      top: top1[0],
      topShare: Math.round(top1Share * 1000) / 10,
      top3Share: Math.round(top3Share * 1000) / 10,
      categories: sorted.length,
      total,
    },
  }
}

/**
 * Analytical findings for a workbook, one sheet at a time.
 *
 * A sheet with no time column still gets concentration; a sheet with no dimension column
 * still gets a trend. A sheet with neither gets nothing — that is the honest answer, not
 * a gap to fill with a weaker claim.
 */
export async function analyzeWorkbook(path: string, sheet?: string): Promise<AnalysisFinding[]> {
  const [semantic, cells] = await Promise.all([
    buildWorkbookSemanticProfile(path, sheet),
    readFile(path).then((data) => readWorkbookCells(data)),
  ])

  const findings: AnalysisFinding[] = []
  for (const sheetProfile of semantic.sheets) {
    const columns: Column[] = sheetProfile.columns
      .filter((column) => column.header !== null && column.role !== 'unknown')
      .map((column) => ({ column: column.column, header: column.header!, role: column.role }))

    const time = columns.find((column) => column.role === 'time')
    const measures = columns.filter((column) => column.role === 'measure')
    const dimensions = columns.filter((column) => column.role === 'dimension')

    const rows = sheetRows(cells, sheetProfile.sheet, columns)

    // A column only counts as a time column if its values actually look like periods.
    // One named 日期 holding free text would otherwise be sorted alphabetically and
    // produce a trend over an order that means nothing.
    const usableTime = time && isPeriodLikeColumn(rows, time) ? time : undefined

    for (const measure of measures) {
      if (usableTime) {
        const trend = trendFinding(rows, usableTime, measure)
        if (trend) findings.push(trend)
      }
      // One dimension is enough. Reporting every measure × dimension pair would bury
      // the one line that matters.
      const dimension = dimensions[0]
      if (dimension) {
        const concentration = concentrationFinding(rows, dimension, measure)
        if (concentration) findings.push(concentration)
      }
    }
  }
  return findings
}

/**
 * True when most of a column's values parse as periods.
 *
 * The header row is skipped. It was not, and the header's own text counted as a
 * non-period: three periods plus a header is 3/4 = 75%, under the 80% bar, so a
 * three-period series was never recognised — while a four-period one passed at exactly
 * 4/5. The bug looked like "short series do not work", which is what sent me looking at
 * the length thresholds instead of here.
 */
function isPeriodLikeColumn(sheet: Sheet, column: Column): boolean {
  let seen = 0
  let periodLike = 0
  for (const [rowNumber, row] of sheet.rows) {
    if (rowNumber === HEADER_ROW) continue
    const label = (row.get(column.column) ?? '').trim()
    if (label === '') continue
    seen++
    if (isPeriodLike(label)) periodLike++
  }
  return seen > 0 && periodLike / seen >= 0.8
}
