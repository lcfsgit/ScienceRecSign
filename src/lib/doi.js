const DOI_SOURCE = String.raw`(?:https?:\/\/(?:dx\.)?doi\.org\/|doi:\s*)?(10\.\d{4,9}\/[^\s,;，；。、<>"']+)`

export function isDoiColumn(col) {
  return [col?.label, col?.headerRaw].some((name) => isDoiName(name))
}

export function isPapersColumn(col) {
  return [col?.label, col?.headerRaw].some((name) => isPapersName(name))
}

export function doiLinks(raw) {
  const text = String(raw ?? '').trim()
  if (!text) return []
  const found = extractDoiLinks(text)
  if (found.length) return found
  return [{ text, href: bingSearch(text) }]
}

export function extractDoiLinks(raw) {
  const text = String(raw ?? '').trim()
  if (!text) return []
  const found = []
  for (const match of text.matchAll(new RegExp(DOI_SOURCE, 'gi'))) {
    const doi = match[1].replace(/[.,);。、]+$/g, '')
    if (found.some((item) => item.text === doi)) continue
    found.push({ text: doi, href: bingSearch(doi) })
  }
  return found
}

export function paperView(raw) {
  const text = String(raw ?? '').trim()
  if (!text) return { kind: 'empty', items: [], html: '', text: '' }
  const data = parseLoose(text)
  const items = data ? collectPapers(data).filter((item) => item.title || item.links.length) : []
  if (items.some((item) => item.links.length)) return { kind: 'items', items, html: '', text: '' }
  if (extractDoiLinks(text).length) return { kind: 'html', items: [], html: linkifyDois(text), text: '' }
  return { kind: 'text', items: [], html: '', text }
}

export function linkifyDois(value) {
  return String(value ?? '').replace(new RegExp(`(<[^>]+>)|(${DOI_SOURCE})`, 'gi'), (all, tag, doi) => {
    if (tag) return tag
    const link = extractDoiLinks(doi)[0]
    if (!link) return all
    return `<a href="${escapeHtml(link.href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(link.text)}</a>`
  })
}

function collectPapers(data) {
  if (typeof data === 'string') {
    const links = extractDoiLinks(data)
    return links.length ? [{ title: '', links }] : []
  }
  if (Array.isArray(data)) return data.flatMap((item) => collectPapers(item))
  if (!data || typeof data !== 'object') return []
  if (Array.isArray(data.papers)) return collectPapers(data.papers)
  const title = String(data.title || data.titleZh || data.titleEn || data.name || data.paper || '')
  const doi = data.doi ?? data.DOI
  if ((doi != null && String(doi).trim()) || title) {
    return [{
      title,
      links: doi == null ? [] : extractDoiLinks(doi)
    }]
  }
  return []
}

function parseLoose(text) {
  let data = text
  for (let i = 0; i < 3 && typeof data === 'string'; i += 1) {
    const trimmed = data.trim()
    if (!trimmed.startsWith('{') && !trimmed.startsWith('[')) return null
    try {
      data = JSON.parse(trimmed)
    } catch {
      return null
    }
  }
  return data && typeof data === 'object' ? data : null
}

function bingSearch(query) {
  return `https://www.bing.com/search?q=${encodeURIComponent(query)}`
}

function isDoiName(name) {
  const text = String(name || '').trim()
  if (!text) return false
  return /(?:^|[^a-z0-9])doi(?:[^a-z0-9]|$)/i.test(text) || /^doi/i.test(text) || /doi$/i.test(text)
}

function isPapersName(name) {
  const text = String(name || '').trim()
  if (!text) return false
  return /(?:^|[^a-z0-9])papers?(?:[^a-z0-9]|$)/i.test(text)
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[ch]))
}
