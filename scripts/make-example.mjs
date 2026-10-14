/**
 * Rebuilds the sample workbook that ships in `examples/`.
 *
 * The file is a front door, so it has to show the one thing this project exists for: a
 * formula that is quietly wrong and that nothing complains about until somebody looks.
 * `金额` is `数量 × 单价` down the column, except for one row where the second reference
 * slipped up a row — `=E9*F8` instead of `=E9*F9`. It still evaluates to a number, it is
 * still plausible, and it is wrong. That is the shape `detectPatternAnomalies` calls a
 * `reference-offset` and `excel_autofix` repairs without a model.
 *
 * A typed-over formula (a `hardcode-break`) would be the more familiar story, but the
 * deterministic repair deliberately does not touch those — it cannot know the intended
 * value, and guessing one would be worse than saying so. The sample demonstrates what the
 * offline half of the loop really does, not what it would be nice for it to do.
 *
 * None of it is real: ten invented orders, three invented regions, three invented
 * products.
 *
 *   node scripts/make-example.mjs
 */
import { mkdirSync } from 'node:fs'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { dirname, join } from 'node:path'
import ExcelJS from 'exceljs'

/**
 * True when this file is the entry point, so importing it (from the demo and the test)
 * does not rewrite the workbook. Comparing `fileURLToPath(import.meta.url)` against the
 * raw `process.argv[1]` does not work on Windows — one is absolute with backslashes, the
 * other is whatever the shell passed, often relative.
 */
function isMain() {
  return process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href
}

const ROOT = fileURLToPath(new URL('..', import.meta.url))

export const EXAMPLE_PATH = join(ROOT, 'examples', 'sales-ledger.xlsx')
export const EXAMPLE_SHEET = '销售台账'
export const HEADERS = ['订单号', '日期', '区域', '产品', '数量', '单价', '金额']

/** Orders in sheet order; `金额` is derived, so it is not listed here. */
export const ORDERS = [
  ['SO-1001', '2026-01-05', '华东', '产品甲', 12, 30],
  ['SO-1002', '2026-01-06', '华北', '产品乙', 8, 45],
  ['SO-1003', '2026-01-07', '华东', '产品丙', 5, 60],
  ['SO-1004', '2026-01-08', '华南', '产品甲', 20, 15],
  ['SO-1005', '2026-01-09', '华北', '产品乙', 3, 200],
  ['SO-1006', '2026-01-12', '华东', '产品丙', 7, 25],
  ['SO-1007', '2026-01-13', '华南', '产品甲', 9, 40],
  ['SO-1008', '2026-01-14', '华北', '产品丙', 4, 55],
  ['SO-1009', '2026-01-15', '华东', '产品乙', 6, 35],
  ['SO-1010', '2026-01-16', '华南', '产品丙', 2, 90],
]

/** The row whose `金额` references the row above instead of its own. */
export const BROKEN_ROW = 9
/** Row 9 should be `=E9*F9`; it says `=E9*F8`. */
export const BROKEN_FORMULA = '=E9*F8'
export const EXPECTED_FORMULA = '=E9*F9'

/** Write the sample workbook and return its path. */
export async function writeExample(path = EXAMPLE_PATH) {
  mkdirSync(dirname(path), { recursive: true })
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'dsh-excel-chat'
  const sheet = workbook.addWorksheet(EXAMPLE_SHEET)
  sheet.addRow(HEADERS)
  sheet.getRow(1).font = { bold: true }

  ORDERS.forEach((order, index) => {
    const row = index + 2
    const amount = row === BROKEN_ROW
      ? { formula: BROKEN_FORMULA.slice(1) }
      : { formula: `E${row}*F${row}` }
    sheet.addRow([...order, amount])
  })
  // A little width so the file is readable when someone opens it to look around.
  sheet.columns.forEach((column) => { column.width = 12 })
  await workbook.xlsx.writeFile(path)
  return path
}

if (isMain()) {
  const path = await writeExample()
  console.log(`wrote ${path}`)
}
