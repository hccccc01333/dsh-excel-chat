const operandSchema = {
  oneOf: [
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        kind: { type: 'string', enum: ['column'], required: true },
        column: { type: 'string', required: true },
      },
    },
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        kind: { type: 'string', enum: ['cell'], required: true },
        cell: { type: 'string', required: true },
      },
    },
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        kind: { type: 'string', enum: ['constant'], required: true },
        value: { type: 'number', required: true },
      },
    },
  ],
} as const

/**
 * dsh tool-DSL schema for FormulaIR. Validates the model's IR payload before
 * `excel_compile_formula` runs.
 */
export const formulaIrSchema = {
  oneOf: [
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        operation: { type: 'string', enum: ['binary'], required: true },
        left: { ...operandSchema, required: true },
        right: { ...operandSchema, required: true },
        operator: { type: 'string', enum: ['+', '-', '*', '/'], required: true },
      },
    },
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        operation: { type: 'string', enum: ['aggregate'], required: true },
        metric: { type: 'string', required: true },
        function: { type: 'string', enum: ['SUMIFS', 'AVERAGEIFS', 'COUNTIFS', 'SUM'], required: true },
        filters: {
          type: 'array',
          required: true,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              column: { type: 'string', required: true },
              value_from: { type: 'string', required: true },
            },
          },
        },
      },
    },
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        operation: { type: 'string', enum: ['ratio'], required: true },
        numerator: { ...operandSchema, required: true },
        denominator: { ...operandSchema, required: true },
      },
    },
  ],
} as const
