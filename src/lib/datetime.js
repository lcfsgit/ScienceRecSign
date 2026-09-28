export function isUpdateDateColumn(col) {
  return [col?.label, col?.headerRaw].some((name) => /^dataSetUpdateDate$/i.test(String(name || '').trim()))
}

export function formatUpdateDate(raw) {
  if (raw == null || String(raw).trim() === '') return ''
  const date = parseUpdateDate(raw)
  if (!date) return String(raw).trim()
  const pad = (value) => String(value).padStart(2, '0')
  return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())} ${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}`
}

function parseUpdateDate(raw) {
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) return raw
  const text = String(raw).trim()
  if (/^-?\d+(?:\.\d+)?$/.test(text)) {
    const value = Number(text)
    if (!Number.isFinite(value)) return null
    const ms = Math.abs(value) < 1e11 ? value * 1000 : value
    const date = new Date(ms)
    return Number.isNaN(date.getTime()) ? null : date
  }
  const iso = text.includes('T') ? text : text.replace(' ', 'T')
  const withZone = /(?:Z|[+-]\d{2}:?\d{2})$/i.test(iso)
    ? iso
    : iso.length === 10
      ? `${iso}T00:00:00Z`
      : `${iso}Z`
  const date = new Date(withZone)
  return Number.isNaN(date.getTime()) ? null : date
}
