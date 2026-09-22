import { buildDemoBuffer } from '../src/lib/demo.js'
import { containsHtml, htmlSource } from '../src/lib/html.js'
import { buildOutputBytes, cellAddress, loadWorkbook, parseHeader } from '../src/lib/sheet.js'

if (!containsHtml('<p>正文<b>重点</b></p>') || containsHtml('a < b')) {
  throw new Error('HTML 识别失败')
}
if (!htmlSource('&lt;p&gt;转义&lt;/p&gt;').includes('<p>')) throw new Error('转义 HTML 未还原')

const header = parseHeader('学科[计算机|生物|物理|化学|医学]')
if (header.label !== '学科' || header.mode !== 'single' || header.options.length !== 5) {
  throw new Error(`表头选项识别失败: ${JSON.stringify(header)}`)
}

const paren = parseHeader('是否入库（是、否、待定）')
if (paren.label !== '是否入库' || paren.options.join(',') !== '是,否,待定') {
  throw new Error(`中文括号选项识别失败: ${JSON.stringify(paren)}`)
}

const buffer = await buildDemoBuffer()
const loaded = await loadWorkbook(buffer, '科学推荐标注示例.xlsx')
const view = loaded.views['标注数据']
if (!view) throw new Error(`缺少标注数据表: ${loaded.sheetNames.join(',')}`)

const byLabel = Object.fromEntries(view.columns.map((col) => [col.label, col]))
const expect = [
  ['学科', 'single', '表头', ['计算机', '生物', '物理', '化学', '医学']],
  ['相关性', 'single', '数据验证', ['高', '中', '低']],
  ['来源类型', 'single', '取值归纳', ['期刊', '会议', '预印本']],
  ['推荐等级', 'single', '选项表', ['强烈推荐', '推荐', '观望', '不推荐']]
]

for (const [label, mode, source, options] of expect) {
  const col = byLabel[label]
  if (!col) throw new Error(`缺少列 ${label}`)
  if (col.mode !== mode || col.optionSource !== source) {
    throw new Error(`${label} 识别为 ${col.mode}/${col.optionSource}`)
  }
  if (options.some((option) => !col.options.includes(option))) {
    throw new Error(`${label} 选项不完整: ${col.options.join('|')}`)
  }
}

if (byLabel['摘要'].mode !== 'text' || byLabel['摘要'].track) {
  throw new Error('摘要应作为正文，不计入完成度')
}
if (!byLabel['备注'].track) throw new Error('备注应计入完成度')
if (!byLabel['打标结果'] || byLabel['打标结果'].role !== 'label-result' || !byLabel['打标结果'].synthetic) {
  throw new Error('数据表应补上打标结果列')
}
if ((loaded.views['选项'].columns || []).some((col) => col.label === '打标结果')) {
  throw new Error('选项表不应新增打标结果')
}
if (!byLabel['摘要'].longText) throw new Error('摘要应识别为大文本')
if (!byLabel['论文ID']?.hidden || byLabel['论文ID'].role !== 'id') throw new Error('论文ID 应隐藏')
if (byLabel['版本号']?.role !== 'version' || !byLabel['版本号'].readonly) throw new Error('版本号应为只读')

const first = view.rows[0]
const edits = {
  标注数据: {
    [`D${first.r + 1}`]: '高',
    [`F${first.r + 1}`]: '强烈推荐',
    [`G${first.r + 1}`]: '优先审阅',
    [cellAddress(first.r, byLabel['打标结果'].col)]: '1'
  }
}

const { bytes, rebuilt } = await buildOutputBytes({
  bookType: loaded.bookType,
  originalBuffer: loaded.originalBuffer,
  views: loaded.views,
  edits,
  sheetPaths: loaded.sheetPaths
})
if (rebuilt) throw new Error('示例文件不应走重建写回')

const again = await loadWorkbook(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength), '科学推荐标注示例.xlsx')
const row = again.views['标注数据'].rows[0]
const againCols = Object.fromEntries(again.views['标注数据'].columns.map((col) => [col.label, col]))
if (row.cells[againCols['相关性'].index] !== '高') throw new Error(`相关性未写回: ${row.cells[againCols['相关性'].index]}`)
if (row.cells[againCols['推荐等级'].index] !== '强烈推荐') throw new Error('推荐等级未写回')
if (row.cells[againCols['备注'].index] !== '优先审阅') throw new Error('备注未写回')
if (row.cells[againCols['标题'].index] !== '图神经网络在药物重定位中的应用') throw new Error('未修改列被改动')
if (againCols['相关性'].options.join(',') !== '高,中,低') throw new Error('写回后数据验证丢失')
if (againCols['推荐等级'].optionSource !== '选项表') throw new Error('写回后选项表丢失')
if (againCols['打标结果'].synthetic) throw new Error('打标结果表头未写入')
if (row.cells[againCols['打标结果'].index] !== '1') throw new Error(`打标结果未写回: ${row.cells[againCols['打标结果'].index]}`)
if (again.views['标注数据'].columns.filter((col) => col.label === '打标结果').length !== 1) {
  throw new Error('打标结果列重复')
}

console.log('roundtrip ok', {
  rows: view.rows.length,
  columns: view.columns.map((col) => `${col.label}:${col.optionSource || col.mode}`)
})
