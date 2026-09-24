import { buildDemoBuffer } from '../src/lib/demo.js'
import { containsHtml, htmlSource } from '../src/lib/html.js'
import { buildOutputBytes, cellAddress, loadWorkbook, parseHeader } from '../src/lib/sheet.js'
import { parseFileTree } from '../src/lib/tree.js'

const tree = parseFileTree('{"file_tree":{"id":"v1","fileName":"V1","path":"/V1","type":"folder","dir":true,"size":0,"children":[{"id":"csv","fileName":"data.csv","path":"/V1/data.csv","type":"file","dir":false,"size":1024}]}}')
if (!tree || tree.fileName !== 'V1' || !tree.dir || tree.children[0]?.fileName !== 'data.csv' || tree.children[0]?.size !== 1024) {
  throw new Error('目录树识别失败')
}
const plain = parseFileTree('{"fileName":"V1","dir":true,"path":"/V1","size":0,"children":[{"fileName":"M3 spectral parameters.xlsx","dir":false,"path":"/V1/M3 spectral parameters.xlsx","size":14855}]}')
if (!plain || plain.fileName !== 'V1' || plain.children[0]?.fileName !== 'M3 spectral parameters.xlsx' || plain.children[0]?.size !== 14855) {
  throw new Error('纯 JSON 目录树识别失败')
}
const soft = parseFileTree("{'fileName': 'V1', 'dir': True, 'children': [{'fileName': 'a.xlsx', 'dir': False, 'size': 12}]}")
if (!soft || soft.fileName !== 'V1' || soft.children[0]?.fileName !== 'a.xlsx') {
  throw new Error('单引号/True False 目录树识别失败')
}
const wrapped = parseFileTree(`<p>{"fileName":"V1","dir":true,"children":[{"fileName":"a.xlsx","dir":false,"size":12}]}</p>`)
if (!wrapped || wrapped.children[0]?.fileName !== 'a.xlsx') throw new Error('包裹的纯 JSON 目录树识别失败')
const doubled = parseFileTree(JSON.stringify('{"fileName":"V1","dir":true,"children":[]}'))
if (!doubled || doubled.fileName !== 'V1') throw new Error('二次编码 JSON 目录树识别失败')
if (parseFileTree('普通备注')) throw new Error('普通文本不应识别为目录树')
if (!containsHtml('<p>正文<b>重点</b></p>') || containsHtml('a < b')) {
  throw new Error('HTML 识别失败')
}
if (!htmlSource('&lt;p&gt;转义&lt;/p&gt;').includes('<p>')) throw new Error('转义 HTML 未还原')

if (typeof DOMParser !== 'undefined') {
  const { sanitizeHtml } = await import('../src/lib/html.js')
  const cleaned = sanitizeHtml('<p style="background-color: rgb(247, 248, 250); color: rgb(102, 102, 102); font-family: SimSun; font-weight:700">正文</p>')
  if (/background-color|color\s*:|font-family/i.test(cleaned)) throw new Error(`应忽略页面样式 style: ${cleaned}`)
  if (!/font-weight:\s*700/i.test(cleaned)) throw new Error(`应保留字重 style: ${cleaned}`)
}

{
  // GBK: 标题,摘要\n中文,测试
  const gbkCsv = Uint8Array.from([
    0xb1, 0xea, 0xcc, 0xe2, 0x2c, 0xd5, 0xaa, 0xd2, 0xaa, 0x0a,
    0xd6, 0xd0, 0xce, 0xc4, 0x2c, 0xb2, 0xe2, 0xca, 0xd4
  ])
  const gbkLoaded = await loadWorkbook(gbkCsv.buffer, 'gbk.csv')
  const gbkView = Object.values(gbkLoaded.views)[0]
  if (gbkView.columns[0]?.label !== '标题' || gbkView.columns[1]?.label !== '摘要') {
    throw new Error(`GBK CSV 表头解码失败: ${gbkView.columns.map((col) => col.label).join(',')}`)
  }
  if (gbkView.rows[0]?.cells?.[0] !== '中文' || gbkView.rows[0]?.cells?.[1] !== '测试') {
    throw new Error(`GBK CSV 内容解码失败: ${gbkView.rows[0]?.cells?.slice(0, 2).join(',')}`)
  }
}

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
if (byLabel['打标结果'] || view.columns.some((col) => col.role === 'label-result' || col.role === 'label-yes')) {
  throw new Error('数据表不应再补总体是否或子维度是否列')
}
if (!byLabel['是否推荐为数据论文'] || byLabel['是否推荐为数据论文'].role !== 'label-paper' || !byLabel['是否推荐为数据论文'].synthetic) {
  throw new Error('数据表应补上是否推荐为数据论文列')
}
for (const key of ['v1', 'v2', 'v3', 'v4', 'v5', 'v6']) {
  const score = view.columns.find((col) => col.role === 'label-score' && col.valueKey === key)
  if (!score?.synthetic) throw new Error(`${key} 价值分档列未补齐`)
  if (score.options.join(',') !== '0,1,2,3,4') throw new Error(`${key}分档选项错误`)
}
if ((loaded.views['选项'].columns || []).some((col) => col.role === 'label-score' || col.role === 'label-paper')) {
  throw new Error('选项表不应新增打标列')
}
if (!byLabel['摘要'].longText) throw new Error('摘要应识别为大文本')
if (!byLabel['论文ID']?.hidden || byLabel['论文ID'].role !== 'id') throw new Error('论文ID 应隐藏')
if (byLabel['版本号']?.role !== 'version' || !byLabel['版本号'].readonly) throw new Error('版本号应为只读')

const first = view.rows[0]
const v1Score = view.columns.find((col) => col.role === 'label-score' && col.valueKey === 'v1')
const edits = {
  标注数据: {
    [`D${first.r + 1}`]: '高',
    [`F${first.r + 1}`]: '强烈推荐',
    [`G${first.r + 1}`]: '优先审阅',
    [cellAddress(first.r, byLabel['是否推荐为数据论文'].col)]: '0',
    [cellAddress(first.r, v1Score.col)]: '3'
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
if (againCols['是否推荐为数据论文']?.synthetic) throw new Error('是否推荐为数据论文表头未写入')
if (row.cells[againCols['是否推荐为数据论文'].index] !== '0') {
  throw new Error(`是否推荐为数据论文未写回: ${row.cells[againCols['是否推荐为数据论文'].index]}`)
}
if (again.views['标注数据'].columns.some((col) => col.role === 'label-result' || col.role === 'label-yes')) {
  throw new Error('写回后不应出现总体是否或子维度是否列')
}
const againV1Score = again.views['标注数据'].columns.find((col) => col.role === 'label-score' && col.valueKey === 'v1')
if (!againV1Score || againV1Score.synthetic) throw new Error('v1分表头未写入')
if (row.cells[againV1Score.index] !== '3') throw new Error(`v1分未写回: ${row.cells[againV1Score.index]}`)

console.log('roundtrip ok', {
  rows: view.rows.length,
  columns: view.columns.map((col) => `${col.label}:${col.optionSource || col.mode}`)
})