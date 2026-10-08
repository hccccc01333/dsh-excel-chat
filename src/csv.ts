/**
 * RFC 4180-ish CSV parsing/writing with configurable delimiters, plus
 * formula-injection guarding for exported cells (borrowed from the
 * noatmark-dsh-plugin idea: neutralize values starting with = + - @).
 */

export function parseCsv(text: string, delimiter = ','): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let inQuotes = false
  let index = 0
  while (index < text.length) {
    const char = text[index]!
    if (inQuotes) {
      if (char === '"') {
        if (text[index + 1] === '"') {
          field += '"'
          index += 2
          continue
        }
        inQuotes = false
        index += 1
        continue
      }
      field += char
      index += 1
      continue
    }
    if (char === '"' && field === '') {
      inQuotes = true
      index += 1
      continue
    }
    if (char === delimiter) {
      row.push(field)
      field = ''
      index += 1
      continue
    }
    if (char === '\n' || char === '\r') {
      if (char === '\r' && text[index + 1] === '\n') index += 1
      row.push(field)
      field = ''
      rows.push(row)
      row = []
      index += 1
      continue
    }
    field += char
    index += 1
  }
  // A trailing record terminator *closes* the last record; it must not open an
  // empty one. Real CSV files end with a newline, so without this guard every
  // import gains a phantom row and stringify→parse stops being stable.
  if (field !== '' || row.length > 0) {
    row.push(field)
    rows.push(row)
  }
  return rows
}

function csvField(value: string, delimiter: string): string {
  if (value.includes(delimiter) || value.includes('"') || value.includes('\n') || value.includes('\r')) {
    return `"${value.replaceAll('"', '""')}"`
  }
  return value
}

/**
 * Neutralize spreadsheet formula injection (=, +, -, @) for literal values.
 * The prefix set follows OWASP CSV Injection guidance: Excel strips a leading
 * tab or carriage return before deciding a cell is a formula, so `\t` and `\r`
 * are attack prefixes too — not just `=`, `+`, `-`, `@`.
 */
export function guardFormulaInjection(value: string): string {
  return /^[=+\-@\t\r]/.test(value) ? `'${value}` : value
}

/**
 * Undo `guardFormulaInjection` for a field read back from a CSV.
 *
 * The guard is one-way unless something reverses it: a text cell `=1+1` was
 * exported as `'=1+1`, and importing that CSV left the apostrophe in the cell,
 * so an export/import round-trip corrupted the value. `guarded` tells the
 * caller the field was text when it was written, which matters — letting the
 * value be re-inferred would turn it back into a live formula.
 */
export function unguardFormulaInjection(value: string): { text: string; guarded: boolean } {
  return /^'[=+\-@\t\r]/.test(value) ? { text: value.slice(1), guarded: true } : { text: value, guarded: false }
}

export function stringifyCsv(rows: string[][], delimiter = ','): string {
  if (rows.length === 0) return ''
  return rows.map((row) => row.map((cell) => csvField(cell, delimiter)).join(delimiter)).join('\r\n') + '\r\n'
}
