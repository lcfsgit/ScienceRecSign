const HTML_TAG = /<\/?(?:a|b|blockquote|br|code|del|div|em|font|h[1-6]|hr|i|img|ins|li|mark|ol|p|pre|s|small|span|strong|sub|sup|table|tbody|td|th|thead|tr|u|ul)\b[^>]*>/i

const ALLOWED = new Set([
  'a', 'b', 'blockquote', 'br', 'code', 'del', 'div', 'em', 'font', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'hr', 'i', 'img', 'ins', 'li', 'mark', 'ol', 'p', 'pre', 's', 'small', 'span', 'strong', 'sub', 'sup',
  'table', 'tbody', 'td', 'th', 'thead', 'tr', 'u', 'ul'
])

const DROP = new Set(['script', 'style', 'iframe', 'object', 'embed', 'link', 'meta', 'noscript', 'svg', 'math', 'form', 'input', 'button', 'textarea'])

const STYLE_PROPS = new Set([
  'background-color', 'font-weight', 'font-style', 'text-decoration', 'text-align',
  'font-size', 'font-family', 'line-height', 'vertical-align'
])

export function containsHtml(value) {
  return HTML_TAG.test(String(value ?? ''))
}

export function htmlSource(value) {
  const raw = String(value ?? '')
  if (containsHtml(raw)) return raw
  if (!/&lt;\s*\/?[a-z]/i.test(raw) && !/&#(?:60|x3c);/i.test(raw)) return ''
  const decoded = decodeEntities(raw)
  return containsHtml(decoded) ? decoded : ''
}

export function sanitizeHtml(value) {
  const source = htmlSource(value)
  if (!source || typeof DOMParser === 'undefined') return ''
  const doc = new DOMParser().parseFromString(source, 'text/html')
  const body = doc.body
  for (const child of [...body.childNodes]) cleanNode(child)
  return body.innerHTML
}

function cleanNode(node) {
  if (node.nodeType === 3) return
  if (node.nodeType !== 1) {
    node.remove()
    return
  }
  const tag = node.tagName.toLowerCase()
  if (DROP.has(tag)) {
    node.remove()
    return
  }
  for (const child of [...node.childNodes]) cleanNode(child)
  if (!ALLOWED.has(tag)) {
    node.replaceWith(...node.childNodes)
    return
  }
  for (const attr of [...node.attributes]) {
    const name = attr.name.toLowerCase()
    if (name.startsWith('on') || name === 'srcdoc' || name === 'formaction') {
      node.removeAttribute(attr.name)
      continue
    }
    if (name === 'style') {
      const style = safeStyle(attr.value)
      if (style) node.setAttribute('style', style)
      else node.removeAttribute(attr.name)
      continue
    }
    if (tag === 'a' && name === 'href') {
      const href = safeUrl(attr.value)
      if (!href) node.removeAttribute(attr.name)
      else {
        node.setAttribute('href', href)
        node.setAttribute('rel', 'noopener noreferrer')
        node.setAttribute('target', '_blank')
      }
      continue
    }
    if (tag === 'img' && name === 'alt') continue
    if (tag === 'img' && name === 'src') {
      const src = safeUrl(attr.value, true)
      if (!src) {
        node.remove()
        return
      }
      node.setAttribute('src', src)
      continue
    }
    if (tag === 'font' && name === 'face') continue
    if ((tag === 'td' || tag === 'th') && (name === 'colspan' || name === 'rowspan')) continue
    node.removeAttribute(attr.name)
  }
  if (tag === 'a' && !node.getAttribute('href')) node.replaceWith(...node.childNodes)
  if (tag === 'img' && !node.getAttribute('src')) node.remove()
}

function safeStyle(value) {
  const parts = []
  for (const chunk of String(value).split(';')) {
    const index = chunk.indexOf(':')
    if (index < 0) continue
    const prop = chunk.slice(0, index).trim().toLowerCase()
    const val = chunk.slice(index + 1).trim()
    if (!STYLE_PROPS.has(prop) || !val || val.length > 80) continue
    if (/url\s*\(|expression|javascript|@import|[\u0000-\u001f\\]/i.test(val)) continue
    parts.push(`${prop}: ${val}`)
  }
  return parts.join('; ')
}

function safeUrl(value, image = false) {
  const text = String(value || '').trim().replace(/[\u0000-\u001f\s]/g, '')
  if (!text || text.length > 2000) return ''
  if (/^(javascript|vbscript|data):/i.test(text)) {
    if (image && /^data:image\/(?:png|jpe?g|gif|webp);base64,/i.test(text)) return text
    return ''
  }
  if (/^https?:\/\//i.test(text)) return text
  if (!image && /^mailto:/i.test(text)) return text
  return ''
}

function decodeEntities(text) {
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
