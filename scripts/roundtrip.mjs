import { formatUpdateDate, isUpdateDateColumn } from '../src/lib/datetime.js'
import { doiLinks, isDoiColumn, isPapersColumn, linkifyDois, paperView } from '../src/lib/doi.js'
import { buildDemoBuffer } from '../src/lib/demo.js'
import { containsHtml, htmlSource } from '../src/lib/html.js'
import { buildOutputBytes, cellAddress, loadWorkbook, parseHeader } from '../src/lib/sheet.js'
import { isRestrictedFileTreeText, parseFileTree, treeLabel } from '../src/lib/tree.js'

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
const blockedName = `\uFFFD`.repeat(12)
const blocked = parseFileTree(`{"fileName":"V1","dir":true,"children":[{"fileName":"${blockedName}.xlsx","dir":false,"size":12},{"fileName":"notes.txt","dir":false,"size":8}]}`)
if (!blocked || treeLabel(blocked.children[0]) !== '该文件限制访问' || treeLabel(blocked.children[1]) !== 'notes.txt') {
  throw new Error('限制访问的文件名未替换')
}
if (!isRestrictedFileTreeText(`{"fileName":"${blockedName}"}`) || isRestrictedFileTreeText('{"fileName":"notes.txt","dir":false}')) {
  throw new Error('无法解析的限制访问目录识别失败')
}
if (!isDoiColumn({ label: 'doi' }) || !isDoiColumn({ headerRaw: '相关DOI' }) || isDoiColumn({ label: '标题' })) {
  throw new Error('DOI 列识别失败')
}
const doi = doiLinks('10.11922/sciencedb.j00104.00101')
if (doi.length !== 1 || doi[0].href !== 'https://www.bing.com/search?q=10.11922%2Fsciencedb.j00104.00101') {
  throw new Error(`DOI 链接错误: ${JSON.stringify(doi)}`)
}
if (doiLinks('https://doi.org/10.1007/s10118-025-3353-3')[0]?.text !== '10.1007/s10118-025-3353-3') {
  throw new Error('完整 DOI 网址未还原成标识')
}
if (doiLinks('').length) throw new Error('空 DOI 不应生成链接')
const papers = paperView('[{"title":"土壤数据集","doi":"10.11922/sciencedb.460"},{"title":"无标识"}]')
if (papers.kind !== 'items' || papers.items[0]?.links[0]?.href !== 'https://www.bing.com/search?q=10.11922%2Fsciencedb.460') {
  throw new Error(`papers 中的 DOI 未变成链接: ${JSON.stringify(papers)}`)
}
if (!papers.items[1] || papers.items[1].links.length) throw new Error('无 DOI 的论文不应生成搜索链接')
const linked = linkifyDois('<p>见 https://doi.org/10.1007/s10118-025-3353-3。</p>')
if (!linked.includes('href="https://www.bing.com/search?q=10.1007%2Fs10118-025-3353-3"') || !linked.startsWith('<p>')) {
  throw new Error(`正文 DOI 未链接化: ${linked}`)
}
if (!isPapersColumn({ label: 'papers' }) || !isPapersColumn({ headerRaw: 'related_papers' }) || isPapersColumn({ label: 'newspapers' })) {
  throw new Error('papers 列识别失败')
}

if (!isUpdateDateColumn({ label: 'dataSetUpdateDate' }) || isUpdateDateColumn({ label: 'doi' })) {
  throw new Error('更新时间列识别失败')
}
const isoStamp = '2022-12-29T08:13:27.948Z'
const formatted = formatUpdateDate(isoStamp)
if (formatted !== '2022-12-29 08:13') throw new Error(`ISO 时间格式化失败: ${formatted}`)
const ms = Date.parse(isoStamp)
if (formatUpdateDate(ms) !== formatted || formatUpdateDate(String(Math.floor(ms / 1000))) !== formatted) {
  throw new Error('时间戳未格式化为同一年月日时分')
}
if (formatUpdateDate('') !== '') throw new Error('空时间应保持空白')

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