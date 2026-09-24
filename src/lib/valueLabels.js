export const VALUE_KEYS = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6']

export const VALUE_LABELS = {
  v1: '问题重要性与资源必要性',
  v2: '相对已有资源的独立增量',
  v3: '再利用潜力与数据产品化',
  v4: '技术可信度与验证',
  v5: '语义互操作与可关联性',
  v6: '应用与生态影响'
}

export const VALUE_RUBRICS = {
  v1: {
    0: '没有明确科学问题、用户或需求。',
    1: '只有宽泛背景，未说明资源为何必要。',
    2: '有明确问题和至少一个用户群，但现有资源缺口证据有限。',
    3: '能够证明现有资源存在客观限制，并且本资源针对这些限制提供改进；至少有一个明确复用需求或用户群。',
    4: '对应重要科学、工程、公共管理、长期监测或基础设施需求；且缺口可核查'
  },
  v2: {
    0: '与已有资源基本相同。',
    1: '仅增加数量、换格式或重新整理，新增价值未证明。',
    2: '在数据覆盖、变量体系、样本规模、计算结果、标签体系、标准化程度、数据质量或来源整合等某个或多个方面存在局部改进。',
    3: '新增内容具有可量化增量，例如新增数据、新标签、新计算结果、新分类体系、跨源融合或标准化映射，并且能够支持至少一个已有资源难以直接完成的任务。',
    4: '形成新模态、新任务、新协议、重要跨源融合或填补明确长期缺口；比较对象、差异字段和新增比例清楚。'
  },
  v3: {
    0: '只能支持原论文中的一个分析。',
    1: '有潜在复用，但字段、层级或使用路径不清楚。',
    2: '至少有一个明确二次分析场景，数据基本可用。数据访问方式、结构化格式、字段设计，可以作为复用能力证据，不要求必须已有用户案例。',
    3: '能够证明存在多个复用方向或用户群；至少提供变量结构、数据字段、示例任务或其他可验证复用依据。',
    4: '可作为基准、参考图谱、训练/验证集、模型校准数据或长期基础设施；文档、访问、限制和版本管理完整。'
  },
  v4: {
    0: '无方法说明或无 QC。',
    1: '有生成流程概述，但无量化验证或误差信息。',
    2: '有 QC、过滤、覆盖度、重复性、一致性或误差指标，但验证链不完整。',
    3: '验证指标量化，异常/缺失处理、误差和限制可由第三方复核。',
    4: '原始—中间—最终数据、处理代码、依赖、验证指标、失败样本/边界情况和敏感性分析形成完整链条。'
  },
  v5: {
    0: '自定义格式，无法与其他资源关联。',
    1: '有基本字段说明，但缺少单位、标准实体 ID 或 schema。',
    2: '采用开放格式，有数据字典和部分标准化字段。',
    3: '本体、标准实体 ID、单位、时间/空间基准和映射表较完整，可与同类资源关联。',
    4: '通用元数据和专业语义均机器可读，提供 schema、映射、验证工具或 API，并有实际关联示例。'
  },
  v6: {
    0: '仅说明潜在用途，无具体资源使用方式。',
    1: '只有作者概括性声称有用。',
    2: '可支持教学、方法开发、局部管理、模型校准或减少重复采集。',
    3: '有真实用户、公共平台、标准、决策、产业或跨机构使用证据。',
    4: '已成为基础设施或被多个独立群体采用，政策、专利、商业使用、下载或复现证据可核查。'
  }
}

export const SCORE_OPTIONS = ['0', '1', '2', '3', '4']

export const PAPER_HEADER = '是否推荐为数据论文'
export const PAPER_LABEL = '是否推荐为数据论文'

export function yesHeader(key) {
  return `${key}是否`
}

export function scoreHeader(key) {
  return `${key}分`
}

export function yesLabel(key) {
  return `是否 · ${VALUE_LABELS[key] || key}`
}

export function scoreLabel(key) {
  return `分档 · ${VALUE_LABELS[key] || key}`
}

export function rubricOf(key, score) {
  const table = VALUE_RUBRICS[key]
  if (!table) return ''
  const n = Number(score)
  return table[n] || ''
}

export function matchValueColumn(header, label) {
  const candidates = [header, label]
    .map((value) => String(value || '').trim())
    .filter(Boolean)
  if (!candidates.length) return null
  for (const text of candidates) {
    const hit = matchOne(text)
    if (hit) return hit
  }
  return null
}

function matchOne(text) {
  if (text === '打标结果') return { key: null, kind: 'result' }
  if (
    text === PAPER_HEADER ||
    text === PAPER_LABEL ||
    text === '推荐数据论文' ||
    text === '数据论文推荐' ||
    text === '是否推荐数据论文'
  ) {
    return { key: null, kind: 'paper' }
  }
  for (const key of VALUE_KEYS) {
    const title = VALUE_LABELS[key]
    if (text === yesHeader(key) || text === yesLabel(key)) return { key, kind: 'yes' }
    if (text === scoreHeader(key) || text === scoreLabel(key)) return { key, kind: 'score' }
    const yesRe = new RegExp(`^(?:${key}[_\\s-]*)?(?:是否)(?:[_\\s·.•-]*${escapeRe(title)})?$`, 'i')
    const scoreRe = new RegExp(`^(?:${key}[_\\s-]*)?(?:分|评分|分档|得分)(?:[_\\s·.•-]*${escapeRe(title)})?$`, 'i')
    const yesAlt = new RegExp(`^是否[_\\s·.•-]*${escapeRe(title)}$`)
    const scoreAlt = new RegExp(`^(?:分|评分|分档|得分)[_\\s·.•-]*${escapeRe(title)}$`)
    const titledYes = new RegExp(`^${escapeRe(title)}[_\\s·.•-]*是否$`)
    const titledScore = new RegExp(`^${escapeRe(title)}[_\\s·.•-]*(?:分|评分|分档|得分)$`)
    if (yesRe.test(text) || yesAlt.test(text) || titledYes.test(text)) return { key, kind: 'yes' }
    if (scoreRe.test(text) || scoreAlt.test(text) || titledScore.test(text)) return { key, kind: 'score' }
  }
  return null
}

export function canonicalYes(raw) {
  const text = String(raw ?? '').trim().toLowerCase()
  if (['1', '是', '是(1)', '是（1）', 'yes', 'y', 'true'].includes(text)) return '1'
  if (['0', '否', '否(0)', '否（0）', 'no', 'n', 'false'].includes(text)) return '0'
  return ''
}

export function canonicalScore(raw) {
  const text = String(raw ?? '').trim()
  if (/^[0-4]$/.test(text)) return text
  return ''
}

function escapeRe(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
