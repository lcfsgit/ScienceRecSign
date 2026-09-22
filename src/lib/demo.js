import * as XLSX from 'xlsx'
import JSZip from 'jszip'

export async function buildDemoBuffer() {
  const headers = ['标题', '摘要', '学科[计算机|生物|物理|化学|医学]', '相关性', '来源类型', '推荐等级', '备注', '论文ID', '版本号']
  const data = [
    ['图神经网络在药物重定位中的应用', '<p>本文提出一种面向<b>药物与靶点</b>关系的图神经网络，用可解释子结构定位潜在适应症。</p><p>并在公开图谱上<span style="color:#d6ff3f">验证召回</span>。</p>', '生物 ', '', '期刊', '', '', 'P2026001', '3.2.0'],
    ['大科学装置的开放共享机制', '围绕大科学装置的机时分配、数据分级与跨机构共享，讨论可执行的开放规则。', '物理', '', '会议', '', '', 'P2026002', '3.2.1'],
    ['可解释学习在临床影像中的进展', '综述临床影像模型的显著性、概念瓶颈与报告生成，比较它们对审稿与临床阅读的帮助。', '医学', '', '期刊', '', '需要补全推荐等级', 'P2026003', '3.1.0'],
    ['高能物理事例筛选的在线触发', '介绍高亮度对撞环境下的触发系统如何在微秒级延迟内保留稀有事例。', '物理', '', '预印本', '', '', 'P2026004', '3.2.0'],
    ['单细胞数据的批次效应校正', '比较多种批次校正策略在细胞图谱整合中的表现，并给出适用边界。', '生物', '', '期刊', '', '', 'P2026005', '3.0.4'],
    ['科研基金评审文本的自动归类', '构建评审意见的主题模型与分类流程，辅助计划处做学科归口。', '计算机', '', '会议', '', '', 'P2026006', '3.2.1']
  ]
  const options = [
    ['字段', '选项'],
    ['推荐等级', '强烈推荐'],
    ['推荐等级', '推荐'],
    ['推荐等级', '观望'],
    ['推荐等级', '不推荐']
  ]

  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([headers, ...data])
  XLSX.utils.book_append_sheet(workbook, sheet, '标注数据')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(options), '选项')
  const raw = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })

  const zip = await JSZip.loadAsync(raw)
  const sheetPath = Object.keys(zip.files).find((name) => /worksheets\/sheet1\.xml$/.test(name))
  let xml = await zip.file(sheetPath).async('string')
  const validation = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="D2:D200" formula1="&quot;高,中,低&quot;"/></dataValidations>'
  if (xml.includes('</sheetData>')) xml = xml.replace('</sheetData>', `</sheetData>${validation}`)
  else xml = xml.replace('</worksheet>', `${validation}</worksheet>`)
  zip.file(sheetPath, xml)
  return zip.generateAsync({ type: 'arraybuffer' })
}
