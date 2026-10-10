import type { ChartInfo, ChartSeries } from './charts.ts'
import { t } from './i18n.ts'
import { expandRange, parseRangeRef } from './charts.ts'

export type ChartAnomalyKind =
  | 'unknown-chart-type'
  | 'no-series'
  | 'missing-categories'
  | 'missing-values'
  | 'invalid-range'
  | 'missing-cells'
  | 'multi-dimensional-range'
  | 'unsorted-dates'

export interface ChartAnomaly {
  kind: ChartAnomalyKind
  chartPath: string
  seriesIndex: number | null
  message: string
}

export interface ChartValidationReport {
  sheetName: string
  chartPath: string
  type: string | null
  seriesCount: number
  anomalies: ChartAnomaly[]
}

const SUPPORTED_TYPES = new Set([
  'barChart',
  'lineChart',
  'pieChart',
  'scatterChart',
  'areaChart',
  'doughnutChart',
  'radarChart',
  'bubbleChart',
])

export function validateCharts(charts: ChartInfo[], cells: Record<string, string>): ChartValidationReport[] {
  const cellIndex = new Map<string, string>()
  for (const [id, value] of Object.entries(cells)) cellIndex.set(id.toUpperCase(), value)
  return charts.map((chart) => {
    const anomalies: ChartAnomaly[] = []
    if (!chart.type) {
      anomalies.push({ kind: 'unknown-chart-type', chartPath: chart.chartPath, seriesIndex: null, message: t('找不到图表类型') })
    } else if (!SUPPORTED_TYPES.has(chart.type)) {
      anomalies.push({ kind: 'unknown-chart-type', chartPath: chart.chartPath, seriesIndex: null, message: t('不支持的图表类型：{type}', { type: chart.type }) })
    }
    if (chart.series.length === 0) {
      anomalies.push({ kind: 'no-series', chartPath: chart.chartPath, seriesIndex: null, message: t('图表没有数据系列') })
    }
    chart.series.forEach((series, index) => {
      validateSeries(series, chart, index, cellIndex, anomalies)
    })
    return {
      sheetName: chart.sheetName,
      chartPath: chart.chartPath,
      type: chart.type,
      seriesCount: chart.series.length,
      anomalies,
    }
  })
}

function validateSeries(
  series: ChartSeries,
  chart: ChartInfo,
  index: number,
  cellIndex: Map<string, string>,
  anomalies: ChartAnomaly[],
): void {
  if (!series.values) {
    anomalies.push({ kind: 'missing-values', chartPath: chart.chartPath, seriesIndex: index, message: t('第 {index} 个系列没有数值引用', { index: index + 1 }) })
  } else {
    validateRange(series.values, chart, index, 'values', cellIndex, anomalies)
  }
  if (!series.categories) {
    anomalies.push({ kind: 'missing-categories', chartPath: chart.chartPath, seriesIndex: index, message: t('第 {index} 个系列没有分类引用', { index: index + 1 }) })
  } else {
    const range = validateRange(series.categories, chart, index, 'categories', cellIndex, anomalies)
    if (range && (chart.type === 'lineChart' || chart.type === 'areaChart')) {
      checkDateOrder(range, chart, index, cellIndex, anomalies)
    }
  }
}

function validateRange(
  ref: string,
  chart: ChartInfo,
  index: number,
  role: string,
  cellIndex: Map<string, string>,
  anomalies: ChartAnomaly[],
): ReturnType<typeof parseRangeRef> {
  const range = parseRangeRef(ref)
  if (!range) {
    anomalies.push({ kind: 'invalid-range', chartPath: chart.chartPath, seriesIndex: index, message: t('第 {index} 个系列（{role}）的区间无效：{ref}', { index: index + 1, role, ref }) })
    return null
  }
  if (range.startColumn !== range.endColumn && range.startRow !== range.endRow) {
    anomalies.push({ kind: 'multi-dimensional-range', chartPath: chart.chartPath, seriesIndex: index, message: t('第 {index} 个系列（{role}）的区间是二维的：{ref}', { index: index + 1, role, ref }) })
  }
  const missing = expandRange(range).filter((id) => !cellIndex.has(id))
  if (missing.length > 0) {
    anomalies.push({
      kind: 'missing-cells',
      chartPath: chart.chartPath,
      seriesIndex: index,
      message: t('第 {index} 个系列（{role}）引用了不存在的单元格：{cells}{more}', { index: index + 1, role, cells: missing.slice(0, 5).join(', '), more: missing.length > 5 ? t('（还有 {count} 个）', { count: missing.length - 5 }) : '' }),
    })
  }
  return range
}

function checkDateOrder(
  range: ReturnType<typeof parseRangeRef>,
  chart: ChartInfo,
  index: number,
  cellIndex: Map<string, string>,
  anomalies: ChartAnomaly[],
): void {
  if (!range) return
  const dates = expandRange(range)
    .map((id) => cellIndex.get(id))
    .filter((value): value is string => value !== undefined)
    .map((value) => Date.parse(value))
    .filter((time) => Number.isFinite(time))
  if (dates.length < 2 || dates.length !== expandRange(range).length) return
  for (let i = 1; i < dates.length; i++) {
    if (dates[i]! < dates[i - 1]!) {
      anomalies.push({ kind: 'unsorted-dates', chartPath: chart.chartPath, seriesIndex: index, message: t('第 {index} 个系列的分类不是日期升序', { index: index + 1 }) })
      return
    }
  }
}
