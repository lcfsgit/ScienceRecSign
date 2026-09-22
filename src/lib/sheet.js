import * as XLSX from 'xlsx'
import JSZip from 'jszip'

const OPTION_SHEET_RE = /^(选项|可选项|选项表|字典|options?|dict|配置)$/i

const KNOWN_SETS = [
  ['是', '否'],
  ['有', '无'],
  ['对', '错'],
  ['通过', '不通过'],
  ['yes', 'no'],
  ['true', 'false'],
  ['Y', 'N']
]

export function bookTypeFromName(name) {
  const ext = String(name || '').split('.').pop()?.toLowerCase()
  if (ext === 'csv') return 'csv'
  if (ext === 'xls') return 'xls'
  if (ext === 'xlsm') return 'xlsm'
  return 'xlsx'
}

export function mimeOf(bookType) {
  if (bookType === 'csv') return 'text/csv;charset=utf-8'
  if (bookType === 'xls') return 'application/vnd.ms-excel'
  if (bookType === 'xlsm') return 'application/vnd.ms-excel.sheet.macroEnabled.12'
  return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
}

export function cellAddress(row, col) {
  return XLSX.utils.encode_cell({ r: row, c: col })
}

export function parseHeader(raw) {
  const text = String(raw || '').trim()
  const multi = text.match(/^(.+?)\s*\{([^}]+)\}\s*$/)
  if (multi) {
    const options = splitOptions(multi[2])
    if (options.length >= 2) return { label: multi[1].trim(), options, mode: 'multi' }
  }
  const single = text.match(/^(.+?)\s*[\[【(（]\s*([^\]】)）]+?)\s*[\]】)）]\s*$/)
  if (single) {
    const options = splitOptions(single[2])
    if (options.length >= 2) return { label: single[1].trim(), options, mode: 'single' }
  }
  return { label: text, options: null, mode: null }
}

export async function loadWorkbook(buffer, filename) {
  const bookType = bookTypeFromName(filename)
  const originalBuffer = buffer.slice(0)
  let workbook
  if (bookType === 'csv') {
    workbook = XLSX.read(decodeCsv(originalBuffer), { type: 'string' })
  } else {
    workbook = XLSX.read(originalBuffer.slice(0), { type: 'array', cellDates: true })
  }
  if (!workbook.SheetNames?.length) {
    throw new Error('空工作簿')
  }

  const zipBacked = bookType === 'xlsx' || bookType === 'xlsm'
  const sheetPaths = zipBacked ? await readSheetPaths(originalBuffer.slice(0)) : {}
  const validations = zipBacked
    ? await readValidations(originalBuffer.slice(0), workbook, sheetPaths)
    : new Map()
  const dictionary = readDictionary(workbook)

  const views = {}
  for (const name of workbook.SheetNames) {
    views[name] = buildView(
      workbook.Sheets[name],
      validations.get(name) || new Map(),
      dictionary,
      !isOptionSheet(name)
    )
  }

  const activeSheet = workbook.SheetNames.find((name) => !isOptionSheet(name)) || workbook.SheetNames[0]
  return {
    bookType,
    originalBuffer,
    sheetPaths,
    sheetNames: workbook.SheetNames.slice(),
    activeSheet,
    views
  }
}

export async function buildOutputBytes({ bookType, originalBuffer, views, edits, sheetPaths }) {
  const outputEdits = withLabelHeaders(views, edits)
  if (bookType === 'csv') {
    const name = Object.keys(views)[0]
    const view = views[name]
    const aoa = [
      view.columns.map((col) => col.headerRaw || col.label),
      ...view.rows.map((row) => row.cells)
    ]
    const csv = '\uFEFF' + XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(aoa))
    return { bytes: new TextEncoder().encode(csv), rebuilt: false }
  }

  if (bookType === 'xls') {
    const wb = XLSX.read(originalBuffer.slice(0), { type: 'array', cellDates: true })
    applyEdits(wb, outputEdits)
    return { bytes: XLSX.write(wb, { bookType: 'xls', type: 'array' }), rebuilt: true }
  }

  const patchable = await sheetsPatchable(originalBuffer.slice(0), sheetPaths, outputEdits)
  if (!patchable) {
    const wb = XLSX.read(originalBuffer.slice(0), { type: 'array', cellDates: true })
    applyEdits(wb, outputEdits)
    return {
      bytes: XLSX.write(wb, { bookType: bookType === 'xlsm' ? 'xlsx' : 'xlsx', type: 'array' }),
      rebuilt: true
    }
  }

  const zip = await JSZip.loadAsync(originalBuffer.slice(0))
  for (const [sheetName, cells] of Object.entries(outputEdits || {})) {
    if (!cells || !Object.keys(cells).length) continue
    const path = sheetPaths[sheetName]
    const file = path && zip.file(path)
    if (!file) {
      const wb = XLSX.read(originalBuffer.slice(0), { type: 'array', cellDates: true })
      applyEdits(wb, outputEdits)
      return { bytes: XLSX.write(wb, { bookType: 'xlsx', type: 'array' }), rebuilt: true }
    }
    const xml = await file.async('string')
    zip.file(path, patchSheetXml(xml, cells))
  }
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  return { bytes, rebuilt: false }
}

function withLabelHeaders(views, edits) {
  const merged = {}
  for (const [sheet, cells] of Object.entries(edits || {})) merged[sheet] = { ...cells }
  for (const [sheetName, view] of Object.entries(views || {})) {
    const sheetEdits = merged[sheetName]
    if (!sheetEdits || !Object.keys(sheetEdits).length || view?.headerRow == null) continue
    for (const col of view.columns || []) {
      if (!col.synthetic) continue
      const addr = cellAddress(view.headerRow, col.col)
      if (sheetEdits[addr] == null) sheetEdits[addr] = col.headerRaw || '打标结果'
    }
  }
  return merged
}

function applyEdits(workbook, edits) {
  for (const [sheetName, cells] of Object.entries(edits || {})) {
    const ws = workbook.Sheets[sheetName]
    if (!ws || !cells) continue
    for (const [addr, value] of Object.entries(cells)) {
      const { r, c } = XLSX.utils.decode_cell(addr)
      writeCell(ws, r, c, value)
    }
  }
}

function writeCell(ws, r, c, value) {
  const addr = XLSX.utils.encode_cell({ r, c })
  if (value == null || value === '') {
    delete ws[addr]
    return
  }
  ws[addr] = { t: 's', v: String(value) }
  const range = ws['!ref']
    ? XLSX.utils.decode_range(ws['!ref'])
    : { s: { r, c }, e: { r, c } }
  range.s.r = Math.min(range.s.r, r)
  range.s.c = Math.min(range.s.c, c)
  range.e.r = Math.max(range.e.r, r)
  range.e.c = Math.max(range.e.c, c)
  ws['!ref'] = XLSX.utils.encode_range(range)
}

function isOptionSheet(name) {
  return OPTION_SHEET_RE.test(String(name || '').trim())
}

function decodeCsv(buffer) {
  const bytes = new Uint8Array(buffer)
  let utf8 = new TextDecoder('utf-8', { fatal: false }).decode(bytes)
  if (utf8.charCodeAt(0) === 0xfeff) utf8 = utf8.slice(1)
  const bad = (utf8.match(/\uFFFD/g) || []).length
  if (!bad) return utf8
  try {
    const gbk = new TextDecoder('gbk').decode(bytes)
    const badGbk = (gbk.match(/\uFFFD/g) || []).length
    if (badGbk < bad) return gbk
  } catch {
    /* 当前环境没有 GBK 解码器时保留 UTF-8 结果 */
  }
  return utf8
}

function cellText(cell) {
  if (!cell || cell.v == null) return ''
  if (typeof cell.v === 'object' && !(cell.v instanceof Date)) {
    if (Array.isArray(cell.v.richText)) return cell.v.richText.map((part) => part.text || '').join('')
    if (cell.v.text) return String(cell.v.text)
  }
  if (typeof cell.w === 'string' && cell.w !== '') return cell.w
  if (cell.v instanceof Date && !Number.isNaN(cell.v.getTime())) {
    const y = cell.v.getFullYear()
    const m = String(cell.v.getMonth() + 1).padStart(2, '0')
    const d = String(cell.v.getDate()).padStart(2, '0')
    return `${y}-${m}-${d}`
  }
  return String(cell.v)
}

function sparseBounds(ws) {
  let minR = Infinity
  let maxR = 0
  let minC = Infinity
  let maxC = 0
  let count = 0
  for (const key of Object.keys(ws)) {
    if (key[0] === '!') continue
    const text = cellText(ws[key]).trim()
    if (!text) continue
    const pos = XLSX.utils.decode_cell(key)
    minR = Math.min(minR, pos.r)
    maxR = Math.max(maxR, pos.r)
    minC = Math.min(minC, pos.c)
    maxC = Math.max(maxC, pos.c)
    count += 1
  }
  if (!count) return null
  return { minR, maxR, minC, maxC }
}

function countNonEmpty(ws, r, minC, maxC) {
  let n = 0
  for (let c = minC; c <= maxC; c += 1) {
    if (cellText(ws[XLSX.utils.encode_cell({ r, c })]).trim()) n += 1
  }
  return n
}

function buildView(ws, validationMap, dictionary, labelColumn = false) {
  if (!ws) return { columns: [], rows: [], banner: '' }
  const bounds = sparseBounds(ws)
  if (!bounds && !validationMap.size) return { columns: [], rows: [], banner: '' }

  let minC = bounds ? bounds.minC : Math.min(...validationMap.keys())
  let maxC = bounds ? bounds.maxC : Math.max(...validationMap.keys())
  let minR = bounds ? bounds.minR : 0
  let maxR = bounds ? bounds.maxR : 0
  for (const col of validationMap.keys()) {
    minC = Math.min(minC, col)
    maxC = Math.max(maxC, col)
  }

  let headerRow = minR
  let banner = ''
  if (bounds && countNonEmpty(ws, minR, minC, maxC) === 1 && minR + 1 <= maxR) {
    const headerCount = countNonEmpty(ws, minR + 1, minC, maxC)
    if (headerCount >= 3) {
      banner = cellText(ws[XLSX.utils.encode_cell({ r: minR, c: minC })]).trim()
      headerRow = minR + 1
    }
  }

  let colIndexes = []
  for (let c = minC; c <= maxC; c += 1) colIndexes.push(c)
  if (colIndexes.length > 80) {
    colIndexes = colIndexes.filter((c) => {
      if (validationMap.has(c)) return true
      if (cellText(ws[XLSX.utils.encode_cell({ r: headerRow, c })]).trim()) return true
      for (let r = headerRow + 1; r <= maxR; r += 1) {
        if (cellText(ws[XLSX.utils.encode_cell({ r, c })]).trim()) return true
      }
      return false
    })
  }

  const headers = colIndexes.map((c) => ({
    col: c,
    letter: XLSX.utils.encode_col(c),
    raw: cellText(ws[XLSX.utils.encode_cell({ r: headerRow, c })]).trim()
  }))

  const rows = []
  for (let r = headerRow + 1; r <= maxR; r += 1) {
    const cells = headers.map((header) => cellText(ws[XLSX.utils.encode_cell({ r, c: header.col })]))
    rows.push({ r, cells, original: cells.slice() })
  }

  const columns = headers.map((header, index) => {
    const values = rows.map((row) => String(row.cells[index] ?? ''))
    return inferColumn(header, index, values, validationMap.get(header.col) || null, dictionary)
  })

  markIdColumns(columns, rows)
  markVersionColumns(columns, rows)
  if (labelColumn) ensureLabelColumn(columns, rows)
  return { columns, rows, banner, headerRow }
}

function markIdColumns(columns, rows) {
  for (const col of columns) {
    if (col.role === 'label-result') continue
    const values = rows.map((row) => String(row.cells[col.index] ?? ''))
    if (!isIdColumn(col, values)) continue
    col.hidden = true
    col.track = false
    col.role = 'id'
    col.longText = false
  }
}

function isIdColumn(col, values) {
  const name = `${col.label || ''} ${col.headerRaw || ''}`.trim()
  if (/(^|[^a-zA-Z0-9])(id|uid|uuid|guid)([^a-zA-Z0-9]|$)/i.test(name)) return true
  if (/(?:^|[_\-\s])id$/i.test(name) || /ID$/.test(name) || /Id$/.test(name)) return true
  if (/(编号|序号|主键)$/.test(col.label || '')) return true
  const nonEmpty = values.map((value) => value.trim()).filter(Boolean)
  if (nonEmpty.length < 2 || new Set(nonEmpty).size !== nonEmpty.length) return false
  const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
  const longNum = /^\d{6,}$/
  const hex = /^[0-9a-f]{16,}$/i
  return nonEmpty.every((value) => uuid.test(value) || longNum.test(value) || hex.test(value))
}

function markVersionColumns(columns, rows) {
  for (const col of columns) {
    if (col.role === 'label-result' || col.hidden) continue
    const values = rows.map((row) => String(row.cells[col.index] ?? ''))
    if (!isVersionColumn(col, values)) continue
    col.role = 'version'
    col.readonly = true
    col.mode = 'text'
    col.track = false
    col.longText = false
    col.optionSource = '只读'
    col.options = []
  }
}

function isVersionColumn(col, values) {
  const name = `${col.label || ''} ${col.headerRaw || ''}`.trim()
  if (/版本号|版本|version|revision/i.test(name) || /(?:^|[_\-\s])ver(?:sion)?$/i.test(name)) return true
  const nonEmpty = values.map((value) => value.trim()).filter(Boolean)
  if (nonEmpty.length < 2) return false
  const version = /^v?\d+(?:\.\d+){1,3}(?:[-_][0-9A-Za-z.]+)?$/i
  return nonEmpty.every((value) => version.test(value))
}

function ensureLabelColumn(columns, rows) {
  const existing = columns.find((col) => col.label === '打标结果' || col.headerRaw === '打标结果')
  if (existing) {
    existing.role = 'label-result'
    existing.mode = 'single'
    existing.options = ['1', '0']
    existing.optionSource = '打标'
    existing.track = true
    existing.synthetic = false
    existing.longText = false
    existing.hidden = false
    return
  }
  const nextCol = columns.reduce((max, col) => Math.max(max, col.col), -1) + 1
  columns.push({
    index: columns.length,
    col: nextCol,
    letter: XLSX.utils.encode_col(nextCol),
    headerRaw: '打标结果',
    label: '打标结果',
    mode: 'single',
    options: ['1', '0'],
    optionSource: '打标',
    separator: '、',
    track: true,
    role: 'label-result',
    synthetic: true,
    longText: false,
    hidden: false
  })
  for (const row of rows) {
    row.cells.push('')
    row.original.push('')
  }
}

function inferColumn(header, index, values, validationOptions, dictionary) {
  const parsed = parseHeader(header.raw)
  const label = parsed.label || `列 ${header.letter}`
  const dictOptions = lookupDict(dictionary, header.raw, label)
  let options = []
  let mode = 'text'
  let source = ''
  let separator = '、'

  if (parsed.options) {
    options = parsed.options
    mode = parsed.mode || 'single'
    source = '表头'
  } else if (validationOptions && validationOptions.length >= 2) {
    options = validationOptions.slice()
    mode = looksMulti(values, validationOptions) ? 'multi' : 'single'
    source = '数据验证'
  } else if (dictOptions && dictOptions.length >= 2) {
    options = dictOptions.slice()
    mode = looksMulti(values, dictOptions) ? 'multi' : 'single'
    source = '选项表'
  } else {
    const multi = inferMulti(values)
    const known = inferKnown(values)
    const single = inferSingle(values)
    if (multi) {
      options = multi.options
      separator = multi.separator
      mode = 'multi'
      source = '取值归纳'
    } else if (known) {
      options = known
      mode = 'single'
      source = '取值归纳'
    } else if (single) {
      options = single
      mode = 'single'
      source = '取值归纳'
    }
  }

  if (mode === 'multi') {
    const detected = detectSeparator(values)
    if (detected) separator = detected
  }

  const trimmed = values.map((value) => value.trim())
  const nonEmpty = trimmed.filter(Boolean)
  const emptyRatio = values.length ? (values.length - nonEmpty.length) / values.length : 1
  const avg = nonEmpty.reduce((sum, value) => sum + value.length, 0) / (nonEmpty.length || 1)
  let track = mode !== 'text' || emptyRatio >= 0.35
  if (mode === 'text' && avg > 80 && emptyRatio < 0.6) track = false
  const longest = values.reduce((max, value) => Math.max(max, String(value).length), 0)
  const longText = mode === 'text' && (longest > 42 || values.some((value) => /[\r\n]/.test(String(value))))

  return {
    index,
    col: header.col,
    letter: header.letter,
    headerRaw: header.raw,
    label,
    mode,
    options,
    optionSource: source,
    separator,
    track,
    role: '',
    synthetic: false,
    longText,
    hidden: false
  }
}

function lookupDict(dictionary, raw, label) {
  if (!dictionary) return null
  if (raw && dictionary.has(raw)) return dictionary.get(raw)
  if (label && dictionary.has(label)) return dictionary.get(label)
  return null
}

function splitOptions(text) {
  return String(text)
    .split(/[|/,，、;；]/)
    .map((part) => part.trim())
    .filter(Boolean)
}

function inferSingle(values) {
  const nonEmpty = values.map((value) => value.trim()).filter(Boolean)
  if (nonEmpty.length < 3) return null
  const uniq = []
  for (const value of nonEmpty) {
    if (!uniq.includes(value)) uniq.push(value)
  }
  if (uniq.length < 2 || uniq.length > 16) return null
  if (uniq.some((value) => value.length > 20)) return null
  const ratio = uniq.length / nonEmpty.length
  if (ratio > 0.6) return null
  return uniq
}

function inferKnown(values) {
  const nonEmpty = values.map((value) => value.trim()).filter(Boolean)
  if (!nonEmpty.length) return null
  const lowered = nonEmpty.map((value) => value.toLowerCase())
  for (const set of KNOWN_SETS) {
    const allowed = new Set(set.map((item) => item.toLowerCase()))
    if (lowered.every((value) => allowed.has(value))) return set.slice()
  }
  return null
}

function inferMulti(values) {
  const nonEmpty = values.map((value) => value.trim()).filter(Boolean)
  if (nonEmpty.length < 2) return null
  const seps = ['|', '、', '，', ',', ';', '；']
  for (const sep of seps) {
    const hits = nonEmpty.filter((value) => value.includes(sep))
    if (hits.length < 2) continue
    if (hits.length / nonEmpty.length < 0.3 && hits.length < nonEmpty.length) continue
    const tokens = []
    for (const value of nonEmpty) {
      for (const part of value.split(sep).map((item) => item.trim()).filter(Boolean)) {
        if (!tokens.includes(part)) tokens.push(part)
      }
    }
    if (tokens.length >= 2 && tokens.length <= 20 && tokens.every((token) => token.length <= 24)) {
      return { options: tokens, separator: sep }
    }
  }
  return null
}

function looksMulti(values, options) {
  const multi = inferMulti(values)
  if (!multi) return false
  return multi.options.every((token) => options.includes(token))
}

function detectSeparator(values) {
  const seps = ['|', '、', '，', ',', ';', '；']
  const score = Object.fromEntries(seps.map((sep) => [sep, 0]))
  for (const value of values) {
    for (const sep of seps) {
      if (String(value).includes(sep)) score[sep] += 1
    }
  }
  const best = seps.slice().sort((a, b) => score[b] - score[a])[0]
  return score[best] > 0 ? best : ''
}

function readDictionary(workbook) {
  const dict = new Map()
  for (const name of workbook.SheetNames) {
    if (!isOptionSheet(name)) continue
    const aoa = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
      header: 1,
      defval: '',
      raw: false,
      blankrows: false
    })
    if (!aoa.length) continue
    const header = (aoa[0] || []).map((value) => String(value ?? '').trim())
    const h0 = header[0] || ''
    const h1 = header[1] || ''
    if (/字段|列名|列|field|name/i.test(h0) && /选项|取值|option|value/i.test(h1)) {
      for (const row of aoa.slice(1)) {
        const field = String(row?.[0] ?? '').trim()
        const cell = String(row?.[1] ?? '').trim()
        if (!field || !cell) continue
        const list = dict.get(field) || []
        for (const part of splitOptions(cell)) {
          if (!list.includes(part)) list.push(part)
        }
        dict.set(field, list)
      }
      continue
    }
    header.forEach((field, idx) => {
      if (!field) return
      const list = dict.get(field) || []
      for (const row of aoa.slice(1)) {
        const value = String(row?.[idx] ?? '').trim()
        if (value && !list.includes(value)) list.push(value)
      }
      if (list.length >= 2) dict.set(field, list)
    })
  }
  return dict
}

function decodeXml(value) {
  return String(value)
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCharCode(parseInt(n, 16)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

function attr(tag, name) {
  const matched = String(tag).match(new RegExp(`(?:^|\\s)${name}="([^"]*)"`))
  return matched ? decodeXml(matched[1]) : ''
}

function textOf(xml, tag) {
  const matched = String(xml).match(new RegExp(`<(?:[\\w]+:)?${tag}\\b[^>]*>([\\s\\S]*?)</(?:[\\w]+:)?${tag}>`))
  if (!matched) return ''
  return decodeXml(matched[1].replace(/<[^>]+>/g, '').trim())
}

function resolveTarget(target) {
  const clean = decodeXml(target).replace(/\\/g, '/').replace(/^\//, '')
  if (clean.startsWith('xl/')) return clean
  return `xl/${clean.replace(/^\.\//, '')}`
}

async function readSheetPaths(buffer) {
  const zip = await JSZip.loadAsync(buffer)
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string')
  if (!workbookXml || !relsXml) return {}
  const rels = new Map()
  for (const tag of relsXml.match(/<Relationship\b[^>]*>/g) || []) {
    const id = attr(tag, 'Id')
    const target = attr(tag, 'Target')
    if (id && target) rels.set(id, target)
  }
  const paths = {}
  for (const tag of workbookXml.match(/<sheet\b[^>]*>/g) || []) {
    const name = attr(tag, 'name')
    const id = attr(tag, 'r:id')
    const target = rels.get(id)
    if (name && target) paths[name] = resolveTarget(target)
  }
  return paths
}

async function readValidations(buffer, workbook, sheetPaths) {
  const zip = await JSZip.loadAsync(buffer)
  const result = new Map()
  for (const [sheetName, path] of Object.entries(sheetPaths)) {
    const file = zip.file(path)
    if (!file) continue
    const xml = await file.async('string')
    const lists = extractListFormulas(xml)
    const map = new Map()
    for (const item of lists) {
      const options = resolveList(item.formula, workbook)
      if (options.length < 2) continue
      for (const col of colsFromSqref(item.sqref)) {
        const prev = map.get(col) || []
        for (const option of options) {
          if (!prev.includes(option)) prev.push(option)
        }
        map.set(col, prev)
      }
    }
    if (map.size) result.set(sheetName, map)
  }
  return result
}

function extractListFormulas(xml) {
  const lists = []
  const classic = /<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/g
  let matched
  while ((matched = classic.exec(xml))) {
    const tag = matched[1] || ''
    const body = matched[2] || ''
    const type = attr(tag, 'type')
    if (type && type !== 'list') continue
    const sqref = attr(tag, 'sqref') || textOf(body, 'sqref')
    const formula = attr(tag, 'formula1') || textOf(body, 'formula1')
    if (sqref && formula) lists.push({ sqref, formula })
  }
  const modern = /<x14:dataValidation\b([^>]*)>([\s\S]*?)<\/x14:dataValidation>/g
  while ((matched = modern.exec(xml))) {
    const type = attr(matched[1], 'type')
    if (type && type !== 'list') continue
    const formula = textOf(matched[2], 'f')
    const sqref = textOf(matched[2], 'sqref')
    if (sqref && formula) lists.push({ sqref, formula })
  }
  return lists
}

function resolveList(formula, workbook) {
  const raw = decodeXml(String(formula).trim())
  if (!raw) return []
  if (raw.startsWith('"')) return parseInlineList(raw)
  const named = workbook.Workbook?.Names?.find((item) => item.Name === raw)
  const ref = named?.Ref || raw
  const fromRange = valuesFromRef(ref, workbook)
  if (fromRange?.length) return fromRange
  if (raw.includes(',')) return parseInlineList(raw.includes('"') ? raw : `"${raw}"`)
  return []
}

function parseInlineList(formula) {
  let text = String(formula).trim()
  if (text.startsWith('"') && text.endsWith('"')) text = text.slice(1, -1)
  text = text.replace(/""/g, '"')
  return text
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean)
}

function valuesFromRef(ref, workbook) {
  const matched = String(ref)
    .trim()
    .match(/^(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i)
  if (!matched) return null
  const sheet = (matched[1] || matched[2]).replace(/''/g, "'")
  const ws = workbook.Sheets[sheet]
  if (!ws) return null
  const c1 = XLSX.utils.decode_col(matched[3])
  const r1 = Number(matched[4]) - 1
  const c2 = XLSX.utils.decode_col(matched[5])
  const r2 = Number(matched[6]) - 1
  const values = []
  const rowStart = Math.min(r1, r2)
  const rowEnd = Math.min(Math.max(r1, r2), rowStart + 500)
  const colStart = Math.min(c1, c2)
  const colEnd = Math.min(Math.max(c1, c2), colStart + 20)
  for (let r = rowStart; r <= rowEnd; r += 1) {
    for (let c = colStart; c <= colEnd; c += 1) {
      const text = cellText(ws[XLSX.utils.encode_cell({ r, c })]).trim()
      if (text && !values.includes(text)) values.push(text)
    }
  }
  return values
}

function colsFromSqref(sqref) {
  const cols = new Set()
  for (const part of String(sqref).split(/\s+/)) {
    if (!part) continue
    const [start, end] = part.replace(/\$/g, '').split(':')
    if (!/^[A-Z]+\d+$/i.test(start)) continue
    const left = XLSX.utils.decode_cell(start)
    const right = end && /^[A-Z]+\d+$/i.test(end) ? XLSX.utils.decode_cell(end) : left
    for (let c = Math.min(left.c, right.c); c <= Math.max(left.c, right.c); c += 1) cols.add(c)
  }
  return [...cols]
}

async function sheetsPatchable(buffer, sheetPaths, edits) {
  const zip = await JSZip.loadAsync(buffer)
  for (const [name, cells] of Object.entries(edits || {})) {
    if (!cells || !Object.keys(cells).length) continue
    const file = sheetPaths?.[name] && zip.file(sheetPaths[name])
    if (!file) return false
    const xml = await file.async('string')
    const tags = xml.match(/<c\b[^>]*>/g) || []
    if (tags.some((tag) => !/\sr="[A-Z]+[0-9]+"/.test(tag))) return false
  }
  return true
}

function patchSheetXml(xml, edits) {
  const grouped = new Map()
  for (const [addr, value] of Object.entries(edits || {})) {
    const pos = XLSX.utils.decode_cell(addr)
    const rowNumber = pos.r + 1
    const list = grouped.get(rowNumber) || []
    list.push({ addr, col: pos.c, value: sanitizeXmlText(value) })
    grouped.set(rowNumber, list)
  }
  let out = xml
  for (const [rowNumber, cells] of grouped) {
    out = patchRow(out, rowNumber, cells)
  }
  return expandDimension(out, edits)
}

function sanitizeXmlText(value) {
  return String(value ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
}

function escapeXmlText(value) {
  return sanitizeXmlText(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function cellXml(addr, value) {
  if (value == null || value === '') return ''
  const space = /^\s|\s$/.test(value) ? ' xml:space="preserve"' : ''
  return `<c r="${addr}" t="inlineStr"><is><t${space}>${escapeXmlText(value)}</t></is></c>`
}

function replaceOnce(haystack, needle, replacement) {
  const index = haystack.indexOf(needle)
  if (index < 0) return haystack
  return haystack.slice(0, index) + replacement + haystack.slice(index + needle.length)
}

function patchRow(xml, rowNumber, cells) {
  const rowRe = new RegExp(`<row\\b[^>]*\\br="${rowNumber}"[^>]*(?:/>|>[\\s\\S]*?</row>)`)
  const found = xml.match(rowRe)
  if (!found) {
    const inner = cells
      .filter((cell) => cell.value !== '')
      .sort((a, b) => a.col - b.col)
      .map((cell) => cellXml(cell.addr, cell.value))
      .join('')
    if (!inner) return xml
    return insertRow(xml, rowNumber, `<row r="${rowNumber}">${inner}</row>`)
  }

  let rowXml = found[0]
  if (rowXml.endsWith('/>')) {
    const inner = cells
      .filter((cell) => cell.value !== '')
      .sort((a, b) => a.col - b.col)
      .map((cell) => cellXml(cell.addr, cell.value))
      .join('')
    const next = inner ? `${rowXml.slice(0, -2)}>${inner}</row>` : rowXml
    return replaceOnce(xml, rowXml, next)
  }

  for (const cell of cells) {
    const cellRe = new RegExp(`<c\\b[^>]*\\br="${cell.addr}"[^>]*(?:/>|>[\\s\\S]*?</c>)`)
    const existing = rowXml.match(cellRe)
    if (existing) {
      rowXml = replaceOnce(rowXml, existing[0], cell.value === '' ? '' : cellXml(cell.addr, cell.value))
    } else if (cell.value !== '') {
      rowXml = insertCell(rowXml, cell.col, cellXml(cell.addr, cell.value))
    }
  }
  return replaceOnce(xml, found[0], rowXml)
}

function insertCell(rowXml, col, piece) {
  const cellRe = /<c\b[^>]*\br="([A-Z]+)[0-9]+"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g
  let matched
  let insertAt = rowXml.lastIndexOf('</row>')
  while ((matched = cellRe.exec(rowXml))) {
    if (XLSX.utils.decode_col(matched[1]) > col) {
      insertAt = matched.index
      break
    }
  }
  if (insertAt < 0) return rowXml + piece
  return rowXml.slice(0, insertAt) + piece + rowXml.slice(insertAt)
}

function insertRow(xml, rowNumber, rowXml) {
  if (xml.includes('<sheetData/>')) {
    return replaceOnce(xml, '<sheetData/>', `<sheetData>${rowXml}</sheetData>`)
  }
  const re = /<row\b[^>]*\br="(\d+)"[^>]*(?:\/>|>[\s\S]*?<\/row>)/g
  let matched
  while ((matched = re.exec(xml))) {
    if (Number(matched[1]) > rowNumber) {
      return xml.slice(0, matched.index) + rowXml + xml.slice(matched.index)
    }
  }
  const end = xml.lastIndexOf('</sheetData>')
  if (end < 0) return xml
  return xml.slice(0, end) + rowXml + xml.slice(end)
}

function expandDimension(xml, edits) {
  const matched = xml.match(/<dimension\b[^>]*\bref="([A-Z]+\d+)(?::([A-Z]+\d+))?"[^>]*\/?>/i)
  if (!matched) return xml
  let start = XLSX.utils.decode_cell(matched[1])
  let end = XLSX.utils.decode_cell(matched[2] || matched[1])
  for (const [addr, value] of Object.entries(edits || {})) {
    if (value === '') continue
    const pos = XLSX.utils.decode_cell(addr)
    start = { r: Math.min(start.r, pos.r), c: Math.min(start.c, pos.c) }
    end = { r: Math.max(end.r, pos.r), c: Math.max(end.c, pos.c) }
  }
  const ref = XLSX.utils.encode_range({ s: start, e: end })
  const next = matched[0].replace(/ref="[^"]+"/i, `ref="${ref}"`)
  return replaceOnce(xml, matched[0], next)
}

export function inferFromValues(values) {
  const multi = inferMulti(values)
  if (multi) return { mode: 'multi', options: multi.options, separator: multi.separator }
  const known = inferKnown(values)
  if (known) return { mode: 'single', options: known, separator: '、' }
  const single = inferSingle(values)
  if (single) return { mode: 'single', options: single, separator: '、' }
  return null
}
