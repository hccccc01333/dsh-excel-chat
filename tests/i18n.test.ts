import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFile, readdir } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { getLanguage, hasTranslation, setLanguage, t } from '../src/i18n.ts'
import { buildWorkbookInsight } from '../src/insight.ts'
import { buildWorkbookMenu } from '../src/menu.ts'
import { writeWorkbookHealthReport } from '../src/health-report.ts'

/**
 * The catalog is keyed by the Chinese source text, so a new message that nobody
 * translated still renders — as Chinese. That is a fine fallback and a terrible
 * silent failure, which is why this test reads every `t('…')` call site out of
 * the source and demands a translation for each one. Without it, adding an
 * English option would quietly leave gaps behind.
 */
test('every t() call site in the source has an English translation', async () => {
  const dir = fileURLToPath(new URL('../src', import.meta.url))
  const files = (await readdir(dir)).filter((name) => name.endsWith('.ts'))
  const missing: string[] = []
  for (const file of files) {
    const source = await readFile(join(dir, file), 'utf8')
    for (const match of source.matchAll(/\bt\(\s*'((?:[^'\\]|\\.)*)'/g)) {
      const key = match[1]!.replace(/\\'/g, "'").replace(/\\\\/g, '\\')
      if (!hasTranslation(key)) missing.push(`${file}: ${key}`)
    }
  }
  assert.deepEqual(missing, [], `untranslated t() keys:\n${missing.join('\n')}`)
})

test('t() substitutes placeholders and falls back to the source text', () => {
  setLanguage('en')
  assert.equal(t('健康分：{score}', { score: 90 }), 'Health score: 90')
  // An unknown key keeps its Chinese wording rather than showing a key name.
  assert.equal(t('这条消息还没有译文'), '这条消息还没有译文')
  setLanguage('zh')
  assert.equal(t('健康分：{score}', { score: 90 }), '健康分：90')
  assert.equal(getLanguage(), 'zh')
})

async function makeWorkbook(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'vera-i18n-'))
  const path = join(dir, 'sales.xlsx')
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('销售')
  sheet.addRow(['区域', '金额'])
  sheet.addRow(['华东', 100])
  sheet.addRow(['华东', null])
  sheet.addRow(['华南', null])
  await workbook.xlsx.writeFile(path)
  return path
}

test('switching to English translates the messages but not the user data', async () => {
  const path = await makeWorkbook()

  setLanguage('zh')
  const zh = await buildWorkbookInsight(path)
  setLanguage('en')
  const en = await buildWorkbookInsight(path)
  setLanguage('zh')

  assert.match(zh.summary, /个工作表/)
  assert.match(en.summary, /worksheet\(s\)/)
  assert.doesNotMatch(en.summary, /[\u4e00-\u9fa5]/, 'the English summary must carry no Chinese of ours')

  // The sheet and column names are the user's own data, so they stay verbatim
  // inside the translated sentence.
  const sheetLine = en.sheets[0]!.summary
  assert.match(sheetLine, /^销售: /)
  assert.match(sheetLine, /金额/)
  assert.doesNotMatch(sheetLine.replace(/销售|金额|区域/g, ''), /[\u4e00-\u9fa5]/)
})

test('the menu and the health report follow the same switch', async () => {
  const path = await makeWorkbook()

  setLanguage('en')
  const menu = await buildWorkbookMenu(path)
  const report = await writeWorkbookHealthReport(path, join(join(path, '..'), 'report.xlsx'))
  setLanguage('zh')

  assert.match(menu.note, /excel_undo/)
  assert.doesNotMatch(menu.note, /[\u4e00-\u9fa5]/)
  assert.ok(menu.suggestions.every((s) => !/[\u4e00-\u9fa5]/.test(s.title)))

  assert.match(report.summary, /^Health score/)
  assert.match(report.summary, /anomalies/)
  // The report sheet name is an identifier written into the file — `validator`
  // skips it by its `_dsh_` prefix and two tool descriptions name it — so it
  // stays the same in every language. Only the surrounding sentence is translated.
  assert.doesNotMatch(report.summary.replace('_dsh_体检报告', ''), /[\u4e00-\u9fa5]/)
  assert.equal(report.reportSheet, '_dsh_体检报告')
})

test('the default stays Chinese so an unconfigured install is unchanged', async () => {
  const path = await makeWorkbook()
  assert.equal(getLanguage(), 'zh')
  const insight = await buildWorkbookInsight(path)
  assert.match(insight.summary, /个工作表/)
})
