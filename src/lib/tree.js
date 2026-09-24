export function looksLikeFileTree(value) {
  const text = String(value ?? '').trim()
  if (text.length < 16) return false
  if (!/file_tree|fileName|"fileName"|'fileName'|fileName\s*:/.test(text) && !/&quot;\s*fileName/.test(text)) {
    return false
  }
  return /children|"dir"|'dir'|dir\s*:|"path"|'path'|path\s*:/.test(text)
}

export function parseFileTree(value) {
  const text = prepareText(value)
  if (!text || !looksLikeFileTree(text)) return null
  for (const candidate of expandCandidates(text)) {
    const data = decodeJson(candidate) ?? extractPayload(candidate)
    const root = pickRoot(data)
    if (root) return root
  }
  return null
}

export function formatSize(size) {
  const n = Number(size)
  if (!Number.isFinite(n) || n < 0) return ''
  if (n < 1024) return `${Math.round(n)} B`
  const units = ['KB', 'MB', 'GB', 'TB']
  let value = n / 1024
  let index = 0
  while (value >= 1024 && index < units.length - 1) {
    value /= 1024
    index += 1
  }
  const digits = value >= 10 ? 0 : 1
  return `${value.toFixed(digits)} ${units[index]}`
}

function prepareText(value) {
  let text = String(value ?? '').trim()
  if (!text) return ''
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1).trim()
  // Excel 文本前缀 '
  if (text.startsWith("'") && (text.includes('"fileName"') || text.includes('file_tree') || text.includes("'fileName'"))) {
    text = text.slice(1).trim()
  }
  // 常见弯引号
  text = text
    .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"')
    .replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'")
  return text
}

function expandCandidates(text) {
  const out = [text]
  const decoded = decodeEntities(text)
  if (decoded !== text) out.push(decoded)
  const unquoted = unwrapQuotes(text)
  if (unquoted !== text) out.push(unquoted)
  if (unquoted !== decoded) {
    const unquotedDecoded = unwrapQuotes(decoded)
    if (!out.includes(unquotedDecoded)) out.push(unquotedDecoded)
  }
  for (const item of [...out]) {
    const noTags = item.replace(/<\/?[a-z][^>]*>/gi, ' ').trim()
    if (noTags && !out.includes(noTags)) out.push(noTags)
    const loose = softenJson(item)
    if (loose && !out.includes(loose)) out.push(loose)
  }
  return out
}

function unwrapQuotes(text) {
  if (text.length < 2) return text
  const quote = text[0]
  if ((quote === '"' || quote === "'") && text[text.length - 1] === quote) {
    const inner = text.slice(1, -1)
    if (quote === '"') return inner.replaceAll('""', '"')
    return inner.replaceAll("''", "'")
  }
  return text
}

function decodeEntities(text) {
  if (!/&(?:lt|gt|quot|apos|amp|#)/i.test(text)) return text
  return text.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|quot|apos|amp);/gi, (all, body) => {
    const key = body.toLowerCase()
    if (key === 'lt') return '<'
    if (key === 'gt') return '>'
    if (key === 'quot') return '"'
    if (key === 'apos') return "'"
    if (key === 'amp') return '&'
    const code = key[1] === 'x' ? Number.parseInt(key.slice(2), 16) : Number.parseInt(key.slice(1), 10)
    if (!Number.isFinite(code) || code <= 0 || code > 0x10ffff) return all
    try {
      return String.fromCodePoint(code)
    } catch {
      return all
    }
  })
}

/** 宽松：单引号键/值、Python True/False/None → 尽量变成可 JSON.parse 的文本 */
function softenJson(text) {
  if (!/[{[]/.test(text)) return ''
  if (!/'/.test(text) && !/\b(?:True|False|None)\b/.test(text)) return ''
  let out = ''
  let inDouble = false
  let inSingle = false
  let escaped = false
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i]
    if (inDouble) {
      out += ch
      if (escaped) escaped = false
      else if (ch === '\\') escaped = true
      else if (ch === '"') inDouble = false
      continue
    }
    if (inSingle) {
      if (escaped) {
        out += ch
        escaped = false
        continue
      }
      if (ch === '\\') {
        out += ch
        escaped = true
        continue
      }
      if (ch === "'") {
        out += '"'
        inSingle = false
        continue
      }
      if (ch === '"') {
        out += '\\"'
        continue
      }
      out += ch
      continue
    }
    if (ch === '"') {
      inDouble = true
      out += ch
      continue
    }
    if (ch === "'") {
      inSingle = true
      out += '"'
      continue
    }
    if (ch === 'T' && text.startsWith('True', i) && isBareWord(text, i, 4)) {
      out += 'true'
      i += 3
      continue
    }
    if (ch === 'F' && text.startsWith('False', i) && isBareWord(text, i, 5)) {
      out += 'false'
      i += 4
      continue
    }
    if (ch === 'N' && text.startsWith('None', i) && isBareWord(text, i, 4)) {
      out += 'null'
      i += 3
      continue
    }
    out += ch
  }
  return out
}

function isBareWord(text, start, len) {
  const before = start === 0 ? '' : text[start - 1]
  const after = text[start + len] || ''
  return !/[A-Za-z0-9_]/.test(before) && !/[A-Za-z0-9_]/.test(after)
}

function decodeJson(text) {
  let data = parseJson(text)
  for (let i = 0; i < 4 && typeof data === 'string'; i += 1) {
    const next = parseJson(data.trim())
    if (next === null) break
    data = next
  }
  return data && typeof data === 'object' ? data : null
}

function pickRoot(data) {
  if (!data || typeof data !== 'object') return null
  if (Array.isArray(data)) {
    const nodes = data.filter(isNode).map(normalize)
    if (nodes.length === 1) return nodes[0]
    if (nodes.length > 1) {
      return { id: '', fileName: '文件', path: '', type: 'folder', dir: true, size: 0, children: nodes }
    }
    return null
  }
  if (typeof data.file_tree === 'string') {
    const nested = pickRoot(decodeJson(data.file_tree) ?? parseJson(data.file_tree))
    if (nested) return nested
  }
  if (isNode(data.file_tree)) return normalize(data.file_tree)
  if (isNode(data)) return normalize(data)
  return null
}

function isNode(node) {
  if (!node || typeof node !== 'object' || Array.isArray(node)) return false
  const name = node.fileName ?? node.path ?? node.name
  if (typeof name !== 'string' || !name.trim()) return false
  return hasDirFlag(node) || node.type === 'folder' || node.type === 'file' || hasChildren(node)
}

function hasDirFlag(node) {
  const dir = node.dir
  return dir === true || dir === false || dir === 1 || dir === 0 || dir === 'true' || dir === 'false' || dir === '1' || dir === '0'
}

function hasChildren(node) {
  if (Array.isArray(node.children)) return true
  if (typeof node.children === 'string' && node.children.trim()) return true
  return false
}

function readChildren(node) {
  if (Array.isArray(node.children)) return node.children
  if (typeof node.children === 'string' && node.children.trim()) {
    const parsed = decodeJson(node.children.trim()) ?? parseJson(node.children.trim())
    if (Array.isArray(parsed)) return parsed
    if (parsed && typeof parsed === 'object') return [parsed]
  }
  return []
}

function normalize(node) {
  const children = readChildren(node).filter(isNode).map(normalize)
  const dirFlag = node.dir
  const dir =
    dirFlag === true ||
    dirFlag === 1 ||
    dirFlag === 'true' ||
    dirFlag === '1' ||
    node.type === 'folder' ||
    (dirFlag !== false && dirFlag !== 0 && dirFlag !== 'false' && dirFlag !== '0' && node.type !== 'file' && children.length > 0)
  const size = Number(node.size)
  const altSize = Number(node['文件大小'] ?? node.fileSize)
  return {
    id: String(node.id ?? ''),
    fileName: String(node.fileName || node.name || node.path || (dir ? '文件夹' : '文件')),
    path: String(node.path || ''),
    type: dir ? 'folder' : 'file',
    dir,
    size: Number.isFinite(size) && size > 0 ? size : Number.isFinite(altSize) ? altSize : 0,
    children
  }
}

function parseJson(text) {
  if (typeof text !== 'string') return text && typeof text === 'object' ? text : null
  try {
    return JSON.parse(text)
  } catch {
    return null
  }
}

function extractPayload(text) {
  // 只从第一个像根对象的 { 或 [ 开始切，避免对每个嵌套括号 O(n²) 扫描
  const start = findRootStart(text)
  if (start < 0) return null
  const sliced = sliceBalanced(text, start)
  if (!sliced) return null
  return parseJson(sliced)
}

function findRootStart(text) {
  let inString = false
  let quote = ''
  let escaped = false
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i]
    if (inString) {
      if (escaped) escaped = false
      else if (ch === '\\') escaped = true
      else if (ch === quote) inString = false
      continue
    }
    if (ch === '"' || ch === "'") {
      inString = true
      quote = ch
      continue
    }
    if (ch === '{' || ch === '[') return i
  }
  return -1
}

function sliceBalanced(text, start) {
  const open = text[start]
  const close = open === '{' ? '}' : ']'
  let depth = 0
  let inString = false
  let quote = ''
  let escaped = false
  for (let i = start; i < text.length; i += 1) {
    const ch = text[i]
    if (inString) {
      if (escaped) escaped = false
      else if (ch === '\\') escaped = true
      else if (ch === quote) inString = false
      continue
    }
    if (ch === '"' || ch === "'") {
      inString = true
      quote = ch
      continue
    }
    if (ch === open) depth += 1
    else if (ch === close) {
      depth -= 1
      if (depth === 0) return text.slice(start, i + 1)
    }
  }
  return null
}
