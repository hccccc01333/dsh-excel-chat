import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { buildCorpusWorkbook } from './helpers.ts'

/**
 * Sheet presentation: protection, comments, hyperlinks, filters, tab colour, grouping,
 * images, tables, validation, page setup, gridlines, zoom, sheet state.
 *
 * The second batch of operations that change no cell. The first (`view.ts`) covered the
 * pane, hidden rows and columns, and widths; these are the rest that ExcelJS can read
 * back after a round trip.
 *
 * That last clause is load-bearing. Sparklines are implemented at the XML layer and
 * ExcelJS reports nothing for them, so an assertion would fail forever and read as a
 * broken operation. Every expectation here was probed against a real round trip before
 * it was written down, and the two kinds that cannot be read this way are deliberately
 * absent rather than quietly passing.
 *
 * Row and column numbers count the header row: row 1 is the header, so the third data
 * row is row 4. Getting that wrong is how the first batch of these tasks failed.
 */
const order = (rows: Array<Array<string | number | null>>) => ({
  name: '订单',
  headers: ['产品', '数量'],
  rows,
})

const three = [['苹果', 10], ['香蕉', 5], ['梨', 7]]

/** A 1x1 PNG, so insertImage has real bytes without a fixture file. */
const DOT_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

export const sheetTasks: FileBenchmarkTask[] = [
  {
    id: 'sheet-protect-and-unprotect',
    category: 'editing',
    name: '保护工作表再取消保护',
    description: '先加保护防止误改，再取消保护恢复可编辑。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-protect-and-unprotect', [order(three)]),
    operations: [
      { op: 'protectSheet', sheet: '订单' },
      { op: 'unprotectSheet', sheet: '订单' },
    ],
    checks: [
      { sheet: '订单', protected: false },
      // Protection changes no cell; a wrong implementation that wiped the data while
      // protecting would pass the state check alone.
      { id: '订单!A2', expect: '苹果' },
    ],
  },
  {
    id: 'sheet-tab-color',
    category: 'editing',
    name: '给工作表标签上色',
    description: '把订单表的标签改成红色，多表时一眼能找到。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-tab-color', [order(three)]),
    operations: [{ op: 'setTabColor', name: '订单', color: 'FFFF0000' }],
    checks: [{ sheet: '订单', tabColor: 'FF0000' }],
  },
  {
    id: 'sheet-autofilter',
    category: 'editing',
    name: '给表头加筛选',
    description: '在表头行开启筛选，方便按产品查看。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-autofilter', [order(three)]),
    operations: [{ op: 'autoFilter', range: '订单!A1:B1' }],
    checks: [{ sheet: '订单', autoFilter: 'A1:B1' }],
  },
  {
    id: 'sheet-cell-comment',
    category: 'editing',
    name: '给单元格加批注',
    description: '在苹果那一格加一条备注，说明这批货的来源。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-cell-comment', [order(three)]),
    operations: [{ op: 'addComment', cell: '订单!A2', text: '产地：山东', author: '库管' }],
    checks: [
      { sheet: '订单', comments: { A2: '产地：山东' } },
      { id: '订单!A2', expect: '苹果' },
    ],
  },
  {
    id: 'sheet-cell-hyperlink',
    category: 'editing',
    name: '给单元格加超链接',
    description: '把产品名链到供应商页面。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-cell-hyperlink', [order(three)]),
    operations: [{ op: 'setHyperlink', cell: '订单!A2', url: 'https://example.com/apple', text: '苹果' }],
    checks: [
      { sheet: '订单', hyperlinks: { A2: 'https://example.com/apple' } },
      // The link must not eat the label: the cell still reads as the product name.
      { id: '订单!A2', expect: '苹果' },
    ],
  },
  {
    id: 'sheet-group-rows',
    category: 'editing',
    name: '按行分组收起明细',
    description: '把订单表最后两行（第 3、4 行）分到一组，方便折叠查看。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-group-rows', [order(three)]),
    operations: [{ op: 'groupRows', sheet: '订单', start: 3, end: 4, level: 1 }],
    checks: [
      { sheet: '订单', outlineRows: { 3: 1, 4: 1 } },
      { id: '订单!A4', expect: '梨' },
    ],
  },
  {
    id: 'sheet-group-columns',
    category: 'editing',
    name: '按列分组',
    description: '把数量列分到一组。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-group-columns', [order(three)]),
    operations: [{ op: 'groupColumns', sheet: '订单', from: 'B', to: 'B', level: 1 }],
    checks: [{ sheet: '订单', outlineColumns: { B: 1 } }],
  },
  {
    id: 'sheet-insert-image',
    category: 'editing',
    name: '插入图片',
    description: '在表格左上角插一个 logo。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-insert-image', [order(three)]),
    operations: [
      { op: 'insertImage', cell: '订单!A1', base64: DOT_PNG_BASE64, width: 16, height: 16 },
    ],
    checks: [
      { sheet: '订单', images: 1 },
      { id: '订单!A1', expect: '产品' },
    ],
  },
  {
    id: 'sheet-add-table',
    category: 'editing',
    name: '把区域变成表格',
    description: '把订单区域变成带条纹的表格，方便筛选和引用。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-add-table', [order(three)]),
    operations: [
      { op: 'addTable', name: '订单表', range: '订单!A1:B4', showRowStripes: true },
    ],
    checks: [
      { sheet: '订单', tables: ['订单表'] },
      { id: '订单!A2', expect: '苹果' },
    ],
  },
  {
    id: 'sheet-data-validation',
    category: 'editing',
    name: '限制数量只能填正整数',
    description: '给数量列加校验，只能填大于 0 的整数。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-data-validation', [order(three)]),
    operations: [
      { op: 'dataValidation', range: '订单!B2:B4', type: 'whole', operator: 'greaterThan', formula1: '0', allowBlank: true },
    ],
    checks: [{ sheet: '订单', dataValidation: { B2: 'whole', B4: 'whole' } }],
  },
  {
    id: 'sheet-page-setup-landscape',
    category: 'editing',
    name: '横向打印并加页眉',
    description: '表格比较宽，改成横向，并在页眉写上表名。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-page-setup-landscape', [order(three)]),
    operations: [
      { op: 'pageSetup', sheet: '订单', orientation: 'landscape', fitToPage: true },
      { op: 'headerFooter', sheet: '订单', oddHeader: '&C订单明细' },
    ],
    checks: [{ sheet: '订单', pageOrientation: 'landscape', oddHeader: '&C订单明细' }],
  },
  {
    id: 'sheet-gridlines-and-zoom',
    category: 'editing',
    name: '关掉网格线并放大',
    description: '截图用，关掉网格线并把视图放大到 150%。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-gridlines-and-zoom', [order(three)]),
    operations: [
      { op: 'showGridLines', sheet: '订单', visible: false },
      { op: 'setZoom', sheet: '订单', zoom: 150 },
    ],
    checks: [{ sheet: '订单', showGridLines: false, zoom: 150 }],
  },
  {
    id: 'sheet-hide-sheet',
    category: 'editing',
    name: '隐藏辅助工作表',
    description: '把中间计算用的表藏起来，只留结果表。',
    buildInput: (dir) => buildCorpusWorkbook(dir, 'sheet-hide-sheet', [
      order(three),
      { name: '辅助', headers: ['说明'], rows: [['中间结果']] },
    ]),
    operations: [{ op: 'hideSheet', name: '辅助', hidden: true }],
    checks: [
      { sheet: '辅助', state: 'hidden' },
      { sheet: '订单', state: 'visible' },
    ],
  },
]
