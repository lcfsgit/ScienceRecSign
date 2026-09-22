<script setup>
import { computed, nextTick, onBeforeUnmount, onMounted, reactive, ref, shallowRef, watch } from 'vue'
import { buildDemoBuffer } from './lib/demo.js'
import { htmlSource, sanitizeHtml } from './lib/html.js'
import { buildOutputBytes, cellAddress, inferFromValues, loadWorkbook, mimeOf } from './lib/sheet.js'

const PICKER_TYPES = [
  {
    description: 'CSV 或 Excel',
    accept: {
      'text/csv': ['.csv'],
      'application/vnd.ms-excel': ['.xls'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm']
    }
  }
]

const canPick = typeof window.showOpenFilePicker === 'function'
const fileInput = ref(null)
const mapRef = ref(null)
const loading = ref(false)
const dragOver = ref(false)
const saving = ref(false)
const toastText = ref('')
const pending = ref(null)
const savedAt = ref('')
const jumpText = ref('1')
const optionDraft = ref('')
const customDrafts = reactive({})
const htmlEdit = reactive({})
const onlyOpen = ref(false)
const activeCol = ref(0)
const rowIndex = ref(0)

const filename = ref('')
const bookType = ref('xlsx')
const fileHandle = shallowRef(null)
const originalBuffer = shallowRef(null)
const sheetPaths = ref({})
const sheetNames = ref([])
const activeSheet = ref('')
const views = ref(null)
const edits = reactive({})

let toastTimer = 0

const view = computed(() => views.value?.[activeSheet.value] || null)
const columns = computed(() => view.value?.columns || [])
const rows = computed(() => view.value?.rows || [])
const formColumns = computed(() => {
  const long = []
  const mark = []
  const rest = []
  for (const col of columns.value) {
    if (col.hidden) continue
    if (col.role === 'label-result') mark.push(col)
    else if (col.mode === 'text' && col.longText) long.push(col)
    else rest.push(col)
  }
  return [...long, ...mark, ...rest]
})
const leadColumns = computed(() => formColumns.value.filter((col) => col.role === 'label-result' || (col.mode === 'text' && col.longText)))
const dataColumns = computed(() => formColumns.value.filter((col) => !leadColumns.value.includes(col)))
const fieldGroups = computed(() => [
  { id: 'lead', cols: leadColumns.value },
  { id: 'data', cols: dataColumns.value }
])
const hiddenColumns = computed(() => columns.value.filter((col) => col.hidden))
const loaded = computed(() => Boolean(views.value))
const activeColumn = computed(() => columns.value[activeCol.value] || null)
const current = computed(() => rows.value[rowIndex.value] || null)

const dirty = computed(() => Object.values(edits).some((cells) => cells && Object.keys(cells).length > 0))
const tracked = computed(() => columns.value.filter((col) => col.track && !col.hidden))

const doneCount = computed(() => {
  if (!tracked.value.length) return 0
  return rows.value.filter((row) => rowDone(row)).length
})

const ratio = computed(() => {
  if (!rows.value.length || !tracked.value.length) return 0
  return doneCount.value / rows.value.length
})

const navList = computed(() => {
  if (!onlyOpen.value) return rows.value.map((_, index) => index)
  return rows.value.map((_, index) => index).filter((index) => !rowDone(rows.value[index]))
})

const canPrev = computed(() => {
  const list = navList.value
  const pos = list.indexOf(rowIndex.value)
  if (pos >= 0) return pos > 0
  return list.some((index) => index < rowIndex.value)
})

const canNext = computed(() => {
  const list = navList.value
  const pos = list.indexOf(rowIndex.value)
  if (pos >= 0) return pos < list.length - 1
  return list.some((index) => index > rowIndex.value)
})

const navLabel = computed(() => {
  if (!onlyOpen.value) return `${pad(rowIndex.value + 1)} / ${pad(rows.value.length)}`
  const pos = navList.value.indexOf(rowIndex.value)
  if (pos < 0) return `未完成 ${navList.value.length}`
  return `未完成 ${pos + 1}/${navList.value.length}`
})

const saveLabel = computed(() => {
  if (saving.value) return '写回中'
  return fileHandle.value ? '写回源文件' : '下载标注结果'
})

const filledCount = computed(() => {
  if (!current.value) return 0
  return tracked.value.filter((col) => String(current.value.cells[col.index] || '').trim()).length
})

const structureIssues = computed(() => {
  const issues = []
  const counts = new Map()
  for (const col of columns.value) {
    if (col.hidden || col.role === 'label-result') continue
    const key = col.label || col.letter
    counts.set(key, (counts.get(key) || 0) + 1)
  }
  for (const [name, count] of counts) {
    if (count > 1) issues.push(`字段「${name}」重复出现`)
  }
  return issues
})

const rowIssues = computed(() => {
  const lines = [...structureIssues.value]
  if (!current.value) return lines
  for (const col of formColumns.value) {
    for (const issue of fieldIssues(col)) lines.push(`${col.label}：${issue}`)
  }
  return lines
})

watch([activeSheet, activeCol], syncDraft)
watch([rowIndex, onlyOpen, activeSheet, loading], () => {
  jumpText.value = String(rowIndex.value + 1)
  for (const key of Object.keys(customDrafts)) delete customDrafts[key]
  for (const key of Object.keys(htmlEdit)) delete htmlEdit[key]
  if (filename.value && rows.value.length) {
    document.title = `${pad(rowIndex.value + 1)} / ${rows.value.length} · 标引`
  }
  nextTick(() => requestAnimationFrame(() => {
    document.querySelectorAll('textarea.full').forEach((el) => sizeText(el))
    drawMap()
  }))
})

onMounted(() => {
  window.addEventListener('keydown', onKey)
  window.addEventListener('beforeunload', onLeave)
  window.addEventListener('resize', drawMap)
})

onBeforeUnmount(() => {
  window.removeEventListener('keydown', onKey)
  window.removeEventListener('beforeunload', onLeave)
  window.removeEventListener('resize', drawMap)
  clearTimeout(toastTimer)
})

function pad(value) {
  return String(value).padStart(2, '0')
}

function rowDone(row) {
  if (!tracked.value.length || !row) return false
  return tracked.value.every((col) => String(row.cells[col.index] || '').trim())
}

function fieldNeed(col) {
  if (!col.track || col.hidden || !current.value) return false
  return !String(current.value.cells[col.index] || '').trim()
}

function fieldIssues(col) {
  if (!current.value || !col || col.hidden) return []
  const text = String(current.value.cells[col.index] ?? '')
  const trimmed = text.trim()
  const issues = []
  if (/^#(N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NULL!|NUM!|GETTING_DATA)/i.test(trimmed)) issues.push('Excel 错误值')
  if (text && text !== trimmed) issues.push('首尾有空白')
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u200B\uFEFF\u00A0]/.test(text)) issues.push('含不可见字符')
  if (/\t/.test(text)) issues.push('含制表符，可能串列')
  if (col.role === 'label-result' && trimmed && !canonicalLabel(text)) issues.push('应是 1 或 0')
  if (col.mode === 'single' && col.role !== 'label-result' && trimmed && col.options.length && !col.options.includes(trimmed)) {
    issues.push('不在可选项内')
  }
  return issues
}

function renderedHtml(col) {
  if (!current.value || !col) return ''
  const source = htmlSource(current.value.cells[col.index])
  return source ? sanitizeHtml(source) : ''
}

function showHtml(col) {
  return Boolean(renderedHtml(col)) && !htmlEdit[col.index]
}

function widgetOf(col) {
  if (col.mode === 'single' || col.mode === 'multi') return col.mode
  const longest = rows.value.reduce((max, row) => Math.max(max, String(row.cells[col.index] || '').length), 0)
  return longest > 42 || col.track ? 'textarea' : 'input'
}

function selectedList(col) {
  const text = String(current.value?.cells[col.index] || '').trim()
  if (!text) return []
  if (col.mode !== 'multi') return [text]
  const sep = col.separator || '、'
  if (text.includes(sep)) return text.split(sep).map((item) => item.trim()).filter(Boolean)
  for (const candidate of ['|', '、', '，', ',', ';', '；']) {
    if (text.includes(candidate)) return text.split(candidate).map((item) => item.trim()).filter(Boolean)
  }
  return [text]
}

function canonicalLabel(raw) {
  const text = String(raw ?? '').trim().toLowerCase()
  if (['1', '是', '是(1)', '是（1）', 'yes', 'y', 'true'].includes(text)) return '1'
  if (['0', '否', '否(0)', '否（0）', 'no', 'n', 'false'].includes(text)) return '0'
  return ''
}

function chipOptions(col) {
  if (col.role === 'label-result') {
    return [
      { label: '是（1）', value: '1' },
      { label: '否（0）', value: '0' }
    ]
  }
  return col.options.map((option) => ({ label: option, value: option }))
}

function isChosen(col, option) {
  if (col.role === 'label-result') return canonicalLabel(current.value?.cells[col.index]) === option.value
  return selectedList(col).includes(option.value)
}

function toast(message) {
  toastText.value = message
  clearTimeout(toastTimer)
  toastTimer = window.setTimeout(() => {
    toastText.value = ''
  }, 2400)
}

function ask(text, run) {
  pending.value = { text, run }
}

function runPending() {
  const job = pending.value?.run
  pending.value = null
  job?.()
}

function syncDraft() {
  const col = columns.value[activeCol.value]
  optionDraft.value = col ? col.options.join(' | ') : ''
}

function clearEdits() {
  for (const key of Object.keys(edits)) delete edits[key]
}

function applyLoaded(data, name, handle) {
  clearEdits()
  filename.value = name
  bookType.value = data.bookType
  fileHandle.value = handle
  originalBuffer.value = data.originalBuffer
  sheetPaths.value = data.sheetPaths
  sheetNames.value = data.sheetNames
  views.value = data.views
  activeSheet.value = data.activeSheet
  activeCol.value = 0
  rowIndex.value = 0
  onlyOpen.value = false
  savedAt.value = ''
  document.title = `${name} · 标引`
  nextTick(() => {
    syncDraft()
    drawMap()
  })
}

function requestOpen() {
  if (dirty.value) {
    ask('当前修改还没写回，更换文件会丢掉这些修改。', chooseFile)
    return
  }
  chooseFile()
}

async function chooseFile() {
  if (canPick) {
    try {
      const [handle] = await window.showOpenFilePicker({ multiple: false, types: PICKER_TYPES })
      const file = await handle.getFile()
      await readFile(file, handle)
      return
    } catch (error) {
      if (error?.name === 'AbortError') return
    }
  }
  fileInput.value?.click()
}

function onInputChange(event) {
  const file = event.target.files?.[0]
  event.target.value = ''
  if (file) readFile(file, null)
}

function onDrop(event) {
  dragOver.value = false
  const file = event.dataTransfer?.files?.[0]
  if (!file || loading.value) return
  if (dirty.value) {
    ask('当前修改还没写回，更换文件会丢掉这些修改。', () => readFile(file, null))
    return
  }
  readFile(file, null)
}

async function readFile(file, handle) {
  const ext = file.name.split('.').pop()?.toLowerCase()
  if (!['csv', 'xls', 'xlsx', 'xlsm'].includes(ext)) {
    toast('请选择 CSV 或 Excel 文件')
    return
  }
  loading.value = true
  try {
    const data = await loadWorkbook(await file.arrayBuffer(), file.name)
    applyLoaded(data, file.name, handle)
    if (!handle) toast('这个文件会在保存时下载，不能直接覆盖原文件')
  } catch (error) {
    console.error(error)
    toast('无法读取这个文件')
  } finally {
    loading.value = false
  }
}

async function loadDemo() {
  if (dirty.value) {
    ask('当前修改还没写回，载入示例会丢掉这些修改。', loadDemoNow)
    return
  }
  await loadDemoNow()
}

async function loadDemoNow() {
  loading.value = true
  try {
    const buffer = await buildDemoBuffer()
    const data = await loadWorkbook(buffer, '科学推荐标注示例.xlsx')
    applyLoaded(data, '科学推荐标注示例.xlsx', null)
  } catch (error) {
    console.error(error)
    toast('示例加载失败')
  } finally {
    loading.value = false
  }
}

function selectSheet(name) {
  if (name === activeSheet.value) return
  flushCustom()
  activeSheet.value = name
  activeCol.value = 0
  rowIndex.value = 0
  onlyOpen.value = false
}

function focusColumn(index) {
  activeCol.value = index
  document.getElementById(`field-${index}`)?.scrollIntoView({ block: 'center' })
}

function setMode(mode) {
  const col = activeColumn.value
  if (!col || col.readonly || col.role === 'version') return
  col.mode = mode
  if (mode !== 'text' && col.options.length < 2) {
    const inferred = inferFromValues(rows.value.map((row) => String(row.cells[col.index] || '')))
    if (inferred) {
      col.options = inferred.options
      col.separator = inferred.separator
      col.optionSource = '取值归纳'
      col.mode = inferred.mode
    }
  }
  if (mode !== 'text') col.track = true
  syncDraft()
  nextTick(drawMap)
}

function toggleTrack() {
  const col = activeColumn.value
  if (!col) return
  col.track = !col.track
  nextTick(drawMap)
}

function onOptionKeydown(event) {
  if (event.isComposing || event.key !== 'Enter') return
  event.preventDefault()
  applyOptions()
}

function applyOptions() {
  const col = activeColumn.value
  if (!col) return
  const options = optionDraft.value
    .split(/[|/,，、;；]/)
    .map((item) => item.trim())
    .filter(Boolean)
  col.options = [...new Set(options)]
  col.optionSource = '自定义'
  if (col.options.length >= 2 && col.mode === 'text') {
    col.mode = 'single'
    col.track = true
  }
  toast(col.options.length ? '已更新可选项' : '已清空可选项')
  nextTick(drawMap)
}

function recognize() {
  const col = activeColumn.value
  if (!col) return
  const inferred = inferFromValues(rows.value.map((row) => String(row.cells[col.index] || '')))
  if (!inferred) {
    toast('这一列没有足够的重复取值')
    return
  }
  col.options = inferred.options
  col.mode = inferred.mode
  col.separator = inferred.separator
  col.optionSource = '取值归纳'
  col.track = true
  syncDraft()
  toast('已按列中取值生成选项')
  nextTick(drawMap)
}

function setCell(colIndex, value) {
  const row = rows.value[rowIndex.value]
  const col = columns.value[colIndex]
  const sheet = activeSheet.value
  if (!row || !col || !sheet || col.readonly || col.role === 'version') return
  row.cells[colIndex] = value
  if (!edits[sheet]) edits[sheet] = {}
  const addr = cellAddress(row.r, col.col)
  if (value === row.original[colIndex]) delete edits[sheet][addr]
  else edits[sheet][addr] = value
  nextTick(drawMap)
}

function chooseSingle(col, option) {
  const raw = String(current.value?.cells[col.index] || '')
  if (col.role === 'label-result') {
    const selected = canonicalLabel(raw)
    if (selected === option.value && (raw === '1' || raw === '0')) {
      setCell(col.index, '')
      return
    }
    setCell(col.index, option.value)
    return
  }
  setCell(col.index, raw === option.value ? '' : option.value)
}

function chooseMulti(col, option) {
  const selected = new Set(selectedList(col))
  if (selected.has(option)) selected.delete(option)
  else selected.add(option)
  const ordered = col.options.filter((item) => selected.has(item))
  for (const extra of selected) {
    if (!ordered.includes(extra)) ordered.push(extra)
  }
  setCell(col.index, ordered.join(col.separator || '、'))
}

function onCustomKeydown(col, event) {
  if (event.isComposing || event.key !== 'Enter') return
  event.preventDefault()
  commitCustom(col)
}

function commitCustom(col) {
  const value = String(customDrafts[col.index] || '').trim()
  if (!value || !col) return
  if (!col.options.includes(value)) col.options.push(value)
  if (col.mode === 'multi') {
    if (!isChosen(col, value)) chooseMulti(col, value)
  } else {
    setCell(col.index, value)
  }
  customDrafts[col.index] = ''
}

function flushCustom() {
  for (const col of columns.value) {
    if (col.mode === 'single' || col.mode === 'multi') commitCustom(col)
  }
}

function autosize(event) {
  sizeText(event.target)
}

function sizeText(el) {
  if (!el?.isConnected) return
  el.style.height = 'auto'
  el.style.height = `${el.scrollHeight + 4}px`
}

function textRows(col) {
  const text = String(current.value?.cells[col.index] || '')
  if (!col.longText) return 3
  let lines = 0
  for (const line of text.split(/\n/)) lines += Math.max(1, Math.ceil(line.length / 32))
  return Math.max(lines, 4)
}

function step(delta) {
  flushCustom()
  const list = navList.value
  if (!list.length) return
  const pos = list.indexOf(rowIndex.value)
  if (pos < 0) {
    if (delta > 0) {
      const later = list.find((index) => index > rowIndex.value)
      if (later != null) rowIndex.value = later
    } else {
      const earlier = [...list].reverse().find((index) => index < rowIndex.value)
      if (earlier != null) rowIndex.value = earlier
    }
    return
  }
  const next = pos + delta
  if (next < 0 || next >= list.length) return
  rowIndex.value = list[next]
}

function jump(value) {
  flushCustom()
  const n = Number(value)
  if (!Number.isInteger(n) || n < 1 || n > rows.value.length) {
    jumpText.value = String(rowIndex.value + 1)
    return
  }
  if (onlyOpen.value && rowDone(rows.value[n - 1])) onlyOpen.value = false
  rowIndex.value = n - 1
}

function onMapClick(event) {
  if (!rows.value.length) return
  const rect = event.currentTarget.getBoundingClientRect()
  const ratioX = (event.clientX - rect.left) / rect.width
  const index = Math.min(rows.value.length - 1, Math.max(0, Math.floor(ratioX * rows.value.length)))
  if (onlyOpen.value && rowDone(rows.value[index])) onlyOpen.value = false
  rowIndex.value = index
}

function drawMap() {
  const canvas = mapRef.value
  if (!canvas) return
  const width = canvas.clientWidth || 260
  const height = 28
  const dpr = Math.min(window.devicePixelRatio || 1, 2)
  canvas.width = Math.floor(width * dpr)
  canvas.height = Math.floor(height * dpr)
  const ctx = canvas.getContext('2d')
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
  ctx.clearRect(0, 0, width, height)
  const list = rows.value
  if (!list.length) return
  const flags = list.map((row) => rowDone(row))
  for (let x = 0; x < width; x += 1) {
    const start = Math.floor((x * list.length) / width)
    const end = Math.max(start + 1, Math.floor(((x + 1) * list.length) / width))
    let open = false
    for (let i = start; i < end && i < list.length; i += 1) {
      if (!flags[i]) open = true
    }
    ctx.fillStyle = open ? '#3c4048' : '#243028'
    ctx.fillRect(x, 8, 1, 12)
  }
  const marker = Math.floor(((rowIndex.value + 0.5) * width) / list.length)
  ctx.fillStyle = '#d6ff3f'
  ctx.fillRect(Math.max(0, marker - 1), 4, 2, 20)
}

function snapshotEdits() {
  const out = {}
  for (const [sheet, cells] of Object.entries(edits)) out[sheet] = { ...cells }
  return out
}

function acceptBaseline(bytes) {
  const copy = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
  originalBuffer.value = copy
  for (const sheetView of Object.values(views.value || {})) {
    for (const row of sheetView.rows) row.original = row.cells.slice()
  }
  clearEdits()
}

function download(blob, name) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = name || 'labeled.xlsx'
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1500)
}

async function save() {
  flushCustom()
  if (!dirty.value || saving.value || !views.value) return
  saving.value = true
  try {
    const { bytes, rebuilt } = await buildOutputBytes({
      bookType: bookType.value,
      originalBuffer: originalBuffer.value,
      views: views.value,
      edits: snapshotEdits(),
      sheetPaths: sheetPaths.value
    })
    const blob = new Blob([bytes], { type: mimeOf(bookType.value === 'xlsm' && rebuilt ? 'xlsx' : bookType.value) })
    let wrote = false
    if (fileHandle.value?.createWritable) {
      try {
        const writable = await fileHandle.value.createWritable()
        await writable.write(blob)
        await writable.close()
        wrote = true
      } catch (error) {
        if (error?.name === 'AbortError') {
          toast('已取消写回')
          return
        }
      }
    }
    if (!wrote) download(blob, filename.value)
    acceptBaseline(bytes)
    const now = new Date()
    savedAt.value = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`
    if (wrote) toast(rebuilt ? '已覆盖源文件，工作簿已重建' : '已写回源文件对应列')
    else toast(rebuilt ? '已下载，工作簿已重建' : '已下载标注结果')
  } catch (error) {
    console.error(error)
    toast('写回失败，请重试')
  } finally {
    saving.value = false
  }
}

function onKey(event) {
  if (event.key === 'Escape' && pending.value) {
    pending.value = null
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
    event.preventDefault()
    save()
    return
  }
  if (!loaded.value || pending.value) return
  const typing = ['INPUT', 'TEXTAREA'].includes(document.activeElement?.tagName)
  if (event.key === 'PageDown' || (event.key === 'ArrowRight' && (event.altKey || !typing))) {
    event.preventDefault()
    step(1)
  } else if (event.key === 'PageUp' || (event.key === 'ArrowLeft' && (event.altKey || !typing))) {
    event.preventDefault()
    step(-1)
  }
}

function onLeave(event) {
  if (!dirty.value) return
  event.preventDefault()
  event.returnValue = ''
}

function modeLabel(mode) {
  if (mode === 'single') return '单选'
  if (mode === 'multi') return '多选'
  return '文本'
}
</script>

<template>
  <div class="app">
    <div class="meter-track" aria-hidden="true">
      <div class="meter" :style="{ transform: `scaleX(${ratio})` }"></div>
    </div>

    <header class="topbar">
      <div class="brand">
        <span class="mark" aria-hidden="true"></span>
        <span class="word">SIGN</span>
        <span class="brand-cn">标引</span>
      </div>
      <div v-if="loaded" class="file-meta">
        <strong class="file-name">{{ filename }}</strong>
        <span class="file-sub">{{ bookType.toUpperCase() }} · {{ doneCount }}/{{ rows.length || 0 }} 行完成</span>
      </div>
      <div class="top-actions">
        <template v-if="loaded">
          <span class="pill" :data-on="dirty">{{ dirty ? '未写回' : (savedAt ? `${savedAt} 已同步` : '已同步') }}</span>
          <button type="button" class="btn ghost" @click="requestOpen">更换文件</button>
          <button type="button" class="btn" :class="{ solid: dirty }" :disabled="!dirty || saving" @click="save">{{ saveLabel }}</button>
        </template>
      </div>
    </header>

    <div v-if="pending" class="confirm">
      <p>{{ pending.text }}</p>
      <button type="button" class="btn ghost" @click="pending = null">取消</button>
      <button type="button" class="btn solid" @click="runPending">继续</button>
    </div>

    <input
      ref="fileInput"
      type="file"
      accept=".csv,.xls,.xlsx,.xlsm,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      hidden
      @change="onInputChange"
    />

    <section
      v-if="!loaded"
      class="gate"
      @dragover.prevent="dragOver = true"
      @dragleave.prevent="dragOver = false"
      @drop.prevent="onDrop"
    >
      <div class="frame" :class="{ over: dragOver }">
        <span class="corner tl"></span>
        <span class="corner tr"></span>
        <span class="corner bl"></span>
        <span class="corner br"></span>
        <p class="eyebrow">SCIENCE REC · SIGN</p>
        <h1>逐行标注</h1>
        <p class="lead">
          选择 CSV 或 Excel。每一页对应源文件中的一行，各列展开成表单；识别到的可选项会同步成单选或多选，填写后写回原来的列。
        </p>
        <div class="gate-actions">
          <button type="button" class="btn solid" :disabled="loading" @click="requestOpen">选择文件</button>
          <button type="button" class="btn ghost" :disabled="loading" @click="loadDemo">载入示例</button>
        </div>
        <ul class="formats">
          <li>CSV</li>
          <li>XLS</li>
          <li>XLSX</li>
          <li>首行作为字段名</li>
        </ul>
        <div v-if="loading" class="loading-line"></div>
        <p class="hint">Chrome / Edge 选择文件后可直接覆盖源文件。拖入或示例文件在保存时下载。</p>
      </div>
    </section>

    <div v-else class="workspace">
      <aside class="rail">
        <div class="rail-head">
          <p class="eyebrow">字段</p>
          <span class="muted">{{ formColumns.length }}</span>
        </div>
        <p v-if="hiddenColumns.length" class="hidden-note">已隐藏 ID 列：{{ hiddenColumns.map((col) => col.label).join('、') }}</p>
        <div class="col-list">
          <button
            v-for="col in formColumns"
            :key="col.letter + col.index"
            type="button"
            class="col-item"
            :class="{ active: activeCol === col.index }"
            @click="focusColumn(col.index)"
          >
            <span class="idx">{{ col.letter }}</span>
            <span class="col-name">{{ col.label }}</span>
            <span class="col-meta">{{ col.role === 'version' ? '只读' : modeLabel(col.mode) }}<template v-if="col.optionSource && col.role !== 'version'"> · {{ col.optionSource }}</template></span>
          </button>
        </div>

        <div v-if="activeColumn?.role === 'version'" class="rail-block">
          <p class="eyebrow">{{ activeColumn.label }}</p>
          <p class="hint">版本号只读，不会改写源文件。</p>
        </div>
        <div v-else-if="activeColumn" class="rail-block">
          <p class="eyebrow">{{ activeColumn.label }}</p>
          <div class="seg" role="group" aria-label="字段类型">
            <button type="button" :class="{ on: activeColumn.mode === 'text' }" @click="setMode('text')">文本</button>
            <button type="button" :class="{ on: activeColumn.mode === 'single' }" @click="setMode('single')">单选</button>
            <button type="button" :class="{ on: activeColumn.mode === 'multi' }" @click="setMode('multi')">多选</button>
          </div>
          <div class="switch-row">
            <span>计入完成度</span>
            <button type="button" class="switch" :class="{ on: activeColumn.track }" :aria-pressed="activeColumn.track" aria-label="计入完成度" @click="toggleTrack">
              <i></i>
            </button>
          </div>
          <label class="editor-label" for="option-draft">可选项</label>
          <input
            id="option-draft"
            v-model="optionDraft"
            class="text-input"
            placeholder="用 | 或逗号分隔"
            @keydown="onOptionKeydown"
          />
          <button type="button" class="mini-btn" @click="applyOptions">应用选项</button>
          <button type="button" class="mini-btn" @click="recognize">按取值重新识别</button>
        </div>

        <div class="rail-foot">
          <div class="map-label">
            <span>行</span>
            <span>{{ rows.length }}</span>
          </div>
          <canvas ref="mapRef" class="minimap" aria-label="按行跳转" @click="onMapClick"></canvas>
          <button type="button" class="filter-btn" :class="{ on: onlyOpen }" @click="onlyOpen = !onlyOpen">
            {{ onlyOpen ? '正在只看未完成' : '只看未完成' }}
          </button>
          <p class="hint">← → 翻页 · Ctrl S 写回</p>
          <p v-if="bookType === 'xls'" class="hint">XLS 写回会重建工作簿，复杂格式可能变化。</p>
          <p v-else-if="fileHandle" class="hint">保存将覆盖源文件中改过的单元格。</p>
          <p v-else class="hint">当前文件没有写入句柄，保存时下载新文件。</p>
        </div>
      </aside>

      <main class="stage">
        <div v-if="sheetNames.length > 1" class="sheet-tabs">
          <button
            v-for="name in sheetNames"
            :key="name"
            type="button"
            :class="{ on: name === activeSheet }"
            @click="selectSheet(name)"
          >
            {{ name }}
          </button>
        </div>

        <div v-if="loading" class="loading-line"></div>

        <div class="stage-scroll">
          <template v-if="current">
            <div class="row-head">
              <h2 class="row-index">{{ pad(rowIndex + 1) }}</h2>
              <div class="row-side">
                <div class="jump-line">
                  <input class="jump" :value="jumpText" inputmode="numeric" aria-label="跳转到数据行" @change="jump($event.target.value)" />
                  <span class="muted">/ {{ rows.length }}</span>
                </div>
                <p class="excel-row">源表第 {{ current.r + 1 }} 行</p>
                <p class="progress-note">
                  <template v-if="tracked.length">本行已填 {{ filledCount }}/{{ tracked.length }}</template>
                  <template v-else>把需要填写的列设为计入完成度</template>
                </p>
              </div>
            </div>
            <p v-if="view?.banner" class="banner">{{ view.banner }}</p>
            <div v-if="rowIssues.length" class="issues">
              <p class="eyebrow">格式</p>
              <ul>
                <li v-for="issue in rowIssues" :key="issue">{{ issue }}</li>
              </ul>
            </div>

            <div v-if="onlyOpen && !navList.length" class="clear-state">
              <p class="eyebrow">CLEAR</p>
              <h2>未完成的行已经没有了</h2>
              <button type="button" class="btn solid" @click="onlyOpen = false">查看全部行</button>
            </div>

            <form v-else :key="activeSheet + '-' + rowIndex" class="form" @submit.prevent>
              <div v-for="group in fieldGroups" :key="group.id" :class="group.id === 'data' ? 'data-grid' : 'lead-stack'">
                <section
                  v-for="col in group.cols"
                  :id="'field-' + col.index"
                  :key="col.index"
                  class="field"
                  :class="{ active: activeCol === col.index, need: fieldNeed(col), reading: col.mode === 'text' && col.longText, result: col.role === 'label-result', version: col.role === 'version', bad: fieldIssues(col).length }"
                  @pointerdown="activeCol = col.index"
                >
                  <div class="field-top">
                    <span class="field-no">{{ pad(col.index + 1) }}</span>
                    <span class="field-label">{{ col.label }}</span>
                    <span v-if="col.role === 'label-result'" class="source">单选 · 写入 1 / 0</span>
                    <span v-else-if="col.role === 'version'" class="source">只读<template v-if="renderedHtml(col)"> · HTML</template></span>
                    <span v-else-if="renderedHtml(col)" class="source">HTML · {{ modeLabel(col.mode) }}</span>
                    <span v-else-if="col.optionSource" class="source">可选项 · {{ col.optionSource }} · {{ col.options.length }}</span>
                    <span v-else class="source">{{ modeLabel(col.mode) }}</span>
                  </div>

                  <div v-if="col.role === 'version'" class="readonly-value">
                    <div v-if="renderedHtml(col)" class="html-view" v-html="renderedHtml(col)"></div>
                    <template v-else>{{ current.cells[col.index] || '—' }}</template>
                  </div>
                  <div v-else-if="col.role !== 'label-result' && showHtml(col)" class="html-block">
                    <div class="html-view" :class="{ full: col.longText }" v-html="renderedHtml(col)"></div>
                    <button type="button" class="html-toggle" @click.stop="htmlEdit[col.index] = true">编辑原文</button>
                  </div>
                  <template v-else-if="col.role !== 'label-result' && widgetOf(col) === 'textarea'">
                    <textarea
                      :class="['grow', { full: col.longText }]"
                      :aria-label="col.label"
                      :value="current.cells[col.index]"
                      :rows="textRows(col)"
                      :ref="(el) => sizeText(el)"
                      @input="setCell(col.index, $event.target.value); autosize($event)"
                    ></textarea>
                    <button v-if="renderedHtml(col)" type="button" class="html-toggle" @click.stop="htmlEdit[col.index] = false">渲染样式</button>
                  </template>
                  <template v-else-if="col.role !== 'label-result' && widgetOf(col) === 'input'">
                    <input
                      class="text-input line-input"
                      :aria-label="col.label"
                      :value="current.cells[col.index]"
                      @input="setCell(col.index, $event.target.value)"
                    />
                    <button v-if="renderedHtml(col)" type="button" class="html-toggle" @click.stop="htmlEdit[col.index] = false">渲染样式</button>
                  </template>
                  <div v-else>
                    <p v-if="col.role !== 'label-result' && !col.options.length" class="empty-options">还没有可选项。可以在左侧补充，或直接输入。</p>
                    <div :class="['chips', { scroll: col.options.length > 12 }]" :role="col.mode === 'multi' ? 'group' : 'radiogroup'" :aria-label="col.label">
                      <button
                        v-for="option in chipOptions(col)"
                        :key="option.value"
                        type="button"
                        class="chip"
                        :class="{ on: isChosen(col, option), wide: col.role === 'label-result' }"
                        :role="col.mode === 'multi' ? 'checkbox' : 'radio'"
                        :aria-checked="isChosen(col, option)"
                        @click="col.mode === 'multi' ? chooseMulti(col, option.value) : chooseSingle(col, option)"
                      >
                        {{ option.label }}
                      </button>
                      <input
                        v-if="col.role !== 'label-result'"
                        class="custom-input"
                        placeholder="其他，回车写入"
                        :value="customDrafts[col.index] || ''"
                        @input="customDrafts[col.index] = $event.target.value"
                        @keydown="onCustomKeydown(col, $event)"
                      />
                    </div>
                  </div>
                  <p v-if="fieldIssues(col).length" class="field-issue">{{ fieldIssues(col).join(' · ') }}</p>
                </section>
              </div>
            </form>
          </template>
          <div v-else class="empty-stage">
            <p class="eyebrow">EMPTY</p>
            <h2>这张表没有可标注的数据行</h2>
            <p class="lead">首行会被当作字段名。换一张工作表，或检查文件是否只有表头。</p>
          </div>
        </div>

        <div class="pager">
          <button type="button" class="btn" :disabled="!canPrev" @click="step(-1)">上一页</button>
          <span class="page-count">{{ navLabel }}</span>
          <button type="button" class="btn" :disabled="!canNext" @click="step(1)">下一页</button>
        </div>
      </main>
    </div>

    <div v-if="toastText" class="toast" role="status">{{ toastText }}</div>
  </div>
</template>
