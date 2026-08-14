import { parseCellId, type ParsedCellId } from './formula.ts'
import type {
  AggregateFormulaIR,
  BinaryFormulaIR,
  ColumnTable,
  FormulaIR,
  FunctionFormulaIR,
  OperandIR,
  RatioFormulaIR,
} from './ir.ts'

export interface CompileContext {
  baseCell: string
  table: ColumnTable
}

const BINARY_OPERATORS = new Set(['+', '-', '*', '/'])
const AGGREGATE_FUNCTIONS = new Set(['SUMIFS', 'AVERAGEIFS', 'COUNTIFS', 'SUM'])
const FUNCTION_NAMES = new Set([
  'VLOOKUP', 'INDEX', 'MATCH', 'ROUND', 'TEXT', 'SUMIF', 'COUNTIF', 'AVERAGE', 'MEDIAN',
  'MAX', 'MIN', 'COUNT', 'COUNTA', 'TODAY', 'YEAR', 'MONTH', 'DAY', 'DATE', 'DATEDIF',
  'EOMONTH', 'SUMIFS', 'AVERAGEIFS', 'COUNTIFS',
])

export function compileFormula(ir: FormulaIR, context: CompileContext): string {
  assertIr(ir)
  const base = parseCellId(context.baseCell)
  return `=${compileNode(ir, base, context)}`
}

function compileNode(ir: FormulaIR, base: ParsedCellId, context: CompileContext): string {
  switch (ir.operation) {
    case 'binary':
      return compileBinary(ir, base, context)
    case 'aggregate':
      return compileAggregate(ir, base, context)
    case 'ratio':
      return compileRatio(ir, base, context)
    case 'function':
      return compileFunction(ir, base, context)
  }
}

function compileBinary(ir: BinaryFormulaIR, base: ParsedCellId, context: CompileContext): string {
  return `${compileOperand(ir.left, base, context)}${ir.operator}${compileOperand(ir.right, base, context)}`
}

function compileRatio(ir: RatioFormulaIR, base: ParsedCellId, context: CompileContext): string {
  return `${compileOperand(ir.numerator, base, context)}/${compileOperand(ir.denominator, base, context)}`
}

function compileFunction(ir: FunctionFormulaIR, base: ParsedCellId, context: CompileContext): string {
  return `${ir.name}(${ir.args.map((arg) => compileOperand(arg, base, context)).join(',')})`
}

function compileAggregate(ir: AggregateFormulaIR, base: ParsedCellId, context: CompileContext): string {
  const metricLetter = columnLetter(ir.metric, context)
  const range = `${context.table.sheet}!$${metricLetter}:$${metricLetter}`
  if (ir.function === 'SUM') return `${ir.function}(${range})`
  const criteria: string[] = []
  for (const filter of ir.filters) {
    const filterLetter = columnLetter(filter.column, context)
    criteria.push(`${context.table.sheet}!$${filterLetter}:$${filterLetter}`, compileFilterValue(filter.value_from, base, context))
  }
  return `${ir.function}(${[range, ...criteria].join(',')})`
}

function compileOperand(operand: OperandIR, base: ParsedCellId, context: CompileContext): string {
  switch (operand.kind) {
    case 'column':
      return `${columnLetter(operand.column, context)}${base.row}`
    case 'cell':
      return operand.cell
    case 'range':
      return operand.range
    case 'constant':
      return String(operand.value)
  }
}

function compileFilterValue(valueFrom: string, base: ParsedCellId, context: CompileContext): string {
  const letter = context.table.columns[valueFrom]
  if (letter) return `${letter}${base.row}`
  if (/^[A-Za-z]{1,3}[0-9]+$/.test(valueFrom)) return valueFrom
  if (/^[+-]?[0-9.]+$/.test(valueFrom)) return valueFrom
  return `"${valueFrom}"`
}

function columnLetter(name: string, context: CompileContext): string {
  const letter = context.table.columns[name]
  if (!letter) throw new Error(`unknown column: ${name}`)
  return letter
}

function assertIr(ir: FormulaIR): void {
  if (!ir || typeof ir !== 'object' || typeof ir.operation !== 'string') {
    throw new Error('invalid Formula IR: operation is required')
  }
  switch (ir.operation) {
    case 'binary': {
      if (!BINARY_OPERATORS.has(ir.operator)) throw new Error(`unsupported binary operator: ${ir.operator}`)
      assertOperand(ir.left)
      assertOperand(ir.right)
      return
    }
    case 'aggregate': {
      if (!AGGREGATE_FUNCTIONS.has(ir.function)) throw new Error(`unsupported aggregate function: ${ir.function}`)
      if (typeof ir.metric !== 'string' || ir.metric.length === 0) throw new Error('aggregate IR requires metric')
      if (!Array.isArray(ir.filters)) throw new Error('aggregate IR requires filters array')
      return
    }
    case 'ratio':
      assertOperand(ir.numerator)
      assertOperand(ir.denominator)
      return
    case 'function': {
      if (!FUNCTION_NAMES.has(ir.name)) throw new Error(`unsupported function: ${ir.name}`)
      if (!Array.isArray(ir.args)) throw new Error('function IR requires args array')
      for (const arg of ir.args) assertOperand(arg)
      return
    }
    default:
      throw new Error(`unsupported IR operation: ${(ir as { operation?: string }).operation}`)
  }
}

function assertOperand(operand: OperandIR): void {
  if (!operand || typeof operand !== 'object' || typeof operand.kind !== 'string') {
    throw new Error('invalid Formula IR: operand kind is required')
  }
  switch (operand.kind) {
    case 'column':
      if (typeof operand.column !== 'string' || operand.column.length === 0) throw new Error('column operand requires column')
      return
    case 'cell':
      if (typeof operand.cell !== 'string' || operand.cell.length === 0) throw new Error('cell operand requires cell')
      return
    case 'range':
      if (typeof operand.range !== 'string' || operand.range.length === 0) throw new Error('range operand requires range')
      return
    case 'constant':
      if (typeof operand.value !== 'number' || !Number.isFinite(operand.value)) throw new Error('constant operand requires finite number')
      return
    default:
      throw new Error(`unsupported operand kind: ${(operand as { kind?: string }).kind}`)
  }
}
