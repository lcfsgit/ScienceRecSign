<script setup>
import { computed, nextTick, onBeforeUnmount, onMounted, reactive, ref, shallowRef, watch } from 'vue'
import FileTree from './components/FileTree.vue'
import { buildDemoBuffer } from './lib/demo.js'
import { htmlSource, sanitizeHtml } from './lib/html.js'
import { buildOutputBytes, cellAddress, loadWorkbook, mimeOf } from './lib/sheet.js'
import { parseFileTree } from './lib/tree.js'
import {
  VALUE_KEYS,
  VALUE_LABELS,
  canonicalScore,
  canonicalYes,
  rubricOf
} from './lib/valueLabels.js'

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
const customDrafts = reactive({})
const htmlEdit = reactive({})
const rubricHover = reactive({})
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
  const rest = []
  for (const col of columns.value) {
    if (col.hidden) continue
    if (isLabelColumn(col)) continue
    if (col.fileTree || col.role === 'file-tree' || fileTreeOf(col) || (col.mode === 'text' && col.longText)) long.push(col)
    else rest.push(col)
  }
  return [...long, ...rest]
})
const labelPaperCol = computed(() => columns.value.find((col) => col.role === 'label-paper') || null)
const valueDimensions = computed(() =>
  VALUE_KEYS.map((key) => ({
    key,
    title: VALUE_LABELS[key],
    score: columns.value.find((col) => col.role === 'label-score' && col.valueKey === key) || null
  })).filter((dim) => dim.score)
)
const leadColumns = computed(() => formColumns.value.filter((col) => col.fileTree || col.role === 'file-tree' || fileTreeOf(col) || (col.mode === 'text' && col.longText)))
const dataColumns = computed(() => formColumns.value.filter((col) => !leadColumns.value.includes(col)))
const fieldGroups = computed(() => [
  { id: 'lead', cols: leadColumns.value },
  { id: 'data', cols: dataColumns.value }
])
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

watch([rowIndex, onlyOpen, activeSheet, loading], () => {
  jumpText.value = String(rowIndex.value + 1)
  for (const key of Object.keys(customDrafts)) delete customDrafts[key]
  for (const key of Object.keys(htmlEdit)) delete htmlEdit[key]
  if (filename.value && rows.value.length) {
    document.title = `${pad(rowIndex.value + 1)} / ${rows.value.length} · 标引`
  }
  nextTick(() => requestAnimationFrame(() => {
    document.querySelectorAll('textarea.full').forEach((el) => {
      if (el.value && el.value.length > 8000 && /fileName|file_tree/.test(el.value)) {
        el.style.height = '8rem'
        return
      }
      sizeText(el)
    })
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

function isLabelColumn(col) {
  return col?.role === 'label-result' || col?.role === 'label-paper' || col?.role === 'label-yes' || col?.role === 'label-score'
}

function isYesNoLabel(col) {
  return col?.role === 'label-result' || col?.role === 'label-paper' || col?.role === 'label-yes'
}

function pad(value) {
  return String(value).padStart(2, '0')
}

function rowDone(row) {
  if (!tracked.value.length || !row) return false
  return tracked.value.every((col) => String(row.cells[col.index] || '').trim())
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
  if (isYesNoLabel(col)) {
    if (trimmed && !canonicalYes(text)) issues.push('应是 1 或 0')
  }
  if (col.role === 'label-score' && trimmed && !canonicalScore(text)) issues.push('应是 0—4')
  if (col.mode === 'single' && !isLabelColumn(col) && trimmed && col.options.length && !col.options.includes(trimmed)) {
    issues.push('不在可选项内')
  }
  return issues
}

const treeCache = new Map()

function renderedHtml(col) {
  if (!current.value || !col) return ''
  if (col.fileTree || col.role === 'file-tree') return ''
  const source = htmlSource(current.value.cells[col.index])
  return source ? sanitizeHtml(source) : ''
}

function fileTreeOf(col) {
  if (!current.value || !col) return null
  const raw = current.value.cells[col.index]
  if (raw == null || raw === '') return null
  const key = typeof raw === 'string' ? raw : String(raw)
  if (treeCache.has(key)) return treeCache.get(key)
  const tree = parseFileTree(key)
  if (treeCache.size > 80) treeCache.clear()
  treeCache.set(key, tree)
  return tree
}

function showHtml(col) {
  return Boolean(renderedHtml(col)) && !fileTreeOf(col) && !htmlEdit[col.index]
}

function showTree(col) {
  return Boolean(fileTreeOf(col)) && !htmlEdit[col.index]
}

function widgetOf(col) {
  if (col.mode === 'single' || col.mode === 'multi') return col.mode
  if (col.fileTree || col.role === 'file-tree') return 'textarea'
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

function chipOptions(col) {
  if (isYesNoLabel(col)) {
    return [
      { label: '是（1）', value: '1' },
      { label: '否（0）', value: '0' }
    ]
  }
  if (col.role === 'label-score') {
    return col.options.map((option) => ({
      label: option,
      value: option,
      rubric: rubricOf(col.valueKey, option)
    }))
  }
  return col.options.map((option) => ({ label: option, value: option }))
}

function isChosen(col, option) {
  if (isYesNoLabel(col)) {
    return canonicalYes(current.value?.cells[col.index]) === option.value
  }
  if (col.role === 'label-score') {
    return canonicalScore(current.value?.cells[col.index]) === option.value
  }
  return selectedList(col).includes(option.value)
}

function dimRubric(dim) {
  if (!dim?.score) return ''
  const hover = rubricHover[dim.key]
  if (hover != null && hover !== '') return rubricOf(dim.key, hover)
  const selected = canonicalScore(current.value?.cells[dim.score.index])
  return selected ? rubricOf(dim.key, selected) : '按 0—4 分档细则给分，悬停或选中可查看说明。'
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

function clearEdits() {
  for (const key of Object.keys(edits)) delete edits[key]
}

function applyLoaded(data, name, handle) {
  clearEdits()
  treeCache.clear()
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
  if (isYesNoLabel(col)) {
    const selected = canonicalYes(raw)
    if (selected === option.value && (raw === '1' || raw === '0')) {
      setCell(col.index, '')
      return
    }
    setCell(col.index, option.value)
    return
  }
  if (col.role === 'label-score') {
    const selected = canonicalScore(raw)
    if (selected === option.value) {
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
      <div class="workspace-top">
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
        <div v-if="loading" class="loading-line workspace-loading"></div>
      </div>

      <div class="workspace-body">
      <aside class="mark-pane">
        <div class="pane-head">
          <div>
            <p class="eyebrow">打标</p>
            <h2 class="pane-title">标注项</h2>
          </div>
          <div v-if="current" class="pane-head-meta">
            <div class="jump-line">
              <input class="jump" :value="jumpText" inputmode="numeric" aria-label="跳转到数据行" @change="jump($event.target.value)" />
              <span class="muted">/ {{ rows.length }}</span>
            </div>
            <p class="progress-note">
              <template v-if="tracked.length">本行 {{ filledCount }}/{{ tracked.length }}</template>
              <template v-else>未设置完成度字段</template>
            </p>
          </div>
        </div>

        <div class="pane-scroll mark-scroll">
          <template v-if="current">
            <div v-if="onlyOpen && !navList.length" class="clear-state compact">
              <p class="eyebrow">CLEAR</p>
              <h2>未完成行已清空</h2>
              <button type="button" class="btn solid" @click="onlyOpen = false">查看全部行</button>
            </div>

            <template v-else-if="labelPaperCol || valueDimensions.length">
            <section
              v-if="labelPaperCol"
              id="field-paper-panel"
              class="field value-panel"
              :class="{ active: activeColumn?.role === 'label-paper' }"
              @pointerdown="activeCol = labelPaperCol.index"
            >
              <div class="field-top">
                <span class="field-no">{{ pad(rowIndex + 1) }}</span>
                <span class="field-label">数据论文</span>
                <span class="source">是否推荐 · 1 / 0</span>
              </div>
              <div class="value-overall">
                <div class="value-dim-head">
                  <strong>是否推荐为数据论文</strong>
                  <span class="source">1 / 0</span>
                </div>
                <div class="chips" role="radiogroup" aria-label="是否推荐为数据论文">
                  <button
                    v-for="option in chipOptions(labelPaperCol)"
                    :key="option.value"
                    type="button"
                    class="chip wide"
                    :class="{ on: isChosen(labelPaperCol, option) }"
                    role="radio"
                    :aria-checked="isChosen(labelPaperCol, option)"
                    @click="chooseSingle(labelPaperCol, option)"
                  >
                    {{ option.label }}
                  </button>
                </div>
              </div>
            </section>

            <section
              v-if="valueDimensions.length"
              id="field-value-panel"
              class="field value-panel"
              :class="{ active: isLabelColumn(activeColumn) && activeColumn?.role !== 'label-paper' }"
            >
              <div class="field-top">
                <span class="field-no">{{ pad(rowIndex + 1) }}</span>
                <span class="field-label">价值打标</span>
                <span class="source">0—4 分档</span>
              </div>

              <div
                v-for="dim in valueDimensions"
                :key="dim.key"
                class="value-dim"
                :class="{ active: activeColumn?.valueKey === dim.key }"
                @pointerdown="activeCol = dim.score.index"
              >
                <div class="value-dim-head">
                  <strong>{{ dim.key.toUpperCase() }} · {{ dim.title }}</strong>
                </div>
                <div class="chips score-chips" role="radiogroup" :aria-label="dim.score.label">
                  <button
                    v-for="option in chipOptions(dim.score)"
                    :key="option.value"
                    type="button"
                    class="chip score"
                    :class="{ on: isChosen(dim.score, option) }"
                    role="radio"
                    :aria-checked="isChosen(dim.score, option)"
                    :title="option.rubric"
                    @mouseenter="rubricHover[dim.key] = option.value"
                    @mouseleave="rubricHover[dim.key] = ''"
                    @focus="rubricHover[dim.key] = option.value"
                    @blur="rubricHover[dim.key] = ''"
                    @click="chooseSingle(dim.score, option)"
                  >
                    {{ option.label }}
                  </button>
                </div>
                <p class="value-rubric">{{ dimRubric(dim) }}</p>
              </div>
            </section>
            </template>

            <div v-else class="empty-stage compact">
              <p class="eyebrow">EMPTY</p>
              <h2>没有可打标列</h2>
            </div>
          </template>
          <div v-else class="empty-stage compact">
            <p class="eyebrow">EMPTY</p>
            <h2>没有数据行</h2>
          </div>
        </div>

        <div class="pane-foot mark-foot">
          <div class="map-label">
            <span>行</span>
            <span>{{ rows.length }}</span>
          </div>
          <canvas ref="mapRef" class="minimap" aria-label="按行跳转" @click="onMapClick"></canvas>
          <button type="button" class="filter-btn" :class="{ on: onlyOpen }" @click="onlyOpen = !onlyOpen">
            {{ onlyOpen ? '正在只看未完成' : '只看未完成' }}
          </button>
          <div class="pager compact">
            <button type="button" class="btn" :disabled="!canPrev" @click="step(-1)">上一页</button>
            <span class="page-count">{{ navLabel }}</span>
            <button type="button" class="btn" :disabled="!canNext" @click="step(1)">下一页</button>
          </div>
          <p class="hint">← → 翻页 · Ctrl S 写回</p>
        </div>
      </aside>

      <main class="detail-pane">
        <div class="pane-head">
          <div>
            <p class="eyebrow">详情</p>
            <h2 class="pane-title">数据字段</h2>
          </div>
          <div v-if="current" class="pane-head-meta">
            <p class="excel-row">源表第 {{ current.r + 1 }} 行</p>
          </div>
        </div>

        <div class="pane-scroll detail-scroll">
          <template v-if="current">
            <p v-if="view?.banner" class="banner">{{ view.banner }}</p>

            <div v-if="onlyOpen && !navList.length" class="clear-state">
              <p class="eyebrow">CLEAR</p>
              <h2>未完成的行已经没有了</h2>
              <button type="button" class="btn solid" @click="onlyOpen = false">查看全部行</button>
            </div>

            <form v-else :key="activeSheet + '-' + rowIndex" class="form detail-form" @submit.prevent>
              <div v-for="group in fieldGroups" :key="group.id" :class="group.id === 'data' ? 'data-grid' : 'lead-stack'">
                <section
                  v-for="col in group.cols"
                  :id="'field-' + col.index"
                  :key="col.index"
                  class="field"
                  :class="{ active: activeCol === col.index, reading: col.mode === 'text' && col.longText, version: col.role === 'version' }"
                  @pointerdown="activeCol = col.index"
                >
                  <div class="field-top">
                    <span class="field-no">{{ col.letter }}</span>
                    <span class="field-label">{{ col.label }}</span>
                    <span v-if="col.role === 'version'" class="source">只读<template v-if="fileTreeOf(col)"> · 目录</template><template v-else-if="renderedHtml(col)"> · HTML</template></span>
                    <span v-else-if="fileTreeOf(col)" class="source">目录</span>
                    <span v-else-if="renderedHtml(col)" class="source">HTML</span>
                  </div>

                  <div v-if="col.role === 'version'" class="readonly-value">
                    <div v-if="fileTreeOf(col)" class="file-tree" role="tree" :aria-label="col.label">
                      <FileTree :node="fileTreeOf(col)" />
                    </div>
                    <div v-else-if="renderedHtml(col)" class="html-view" v-html="renderedHtml(col)"></div>
                    <template v-else>{{ current.cells[col.index] || '—' }}</template>
                  </div>
                  <div v-else-if="showTree(col)" class="html-block">
                    <div class="file-tree" role="tree" :aria-label="col.label">
                      <FileTree :node="fileTreeOf(col)" />
                    </div>
                    <button type="button" class="html-toggle" @click.stop="htmlEdit[col.index] = true">编辑原文</button>
                  </div>
                  <div v-else-if="showHtml(col)" class="html-block">
                    <div class="html-view" :class="{ full: col.longText }" v-html="renderedHtml(col)"></div>
                    <button type="button" class="html-toggle" @click.stop="htmlEdit[col.index] = true">编辑原文</button>
                  </div>
                  <template v-else-if="widgetOf(col) === 'textarea'">
                    <textarea
                      :class="['grow', { full: col.longText }]"
                      :aria-label="col.label"
                      :value="current.cells[col.index]"
                      :rows="textRows(col)"
                      :ref="(el) => sizeText(el)"
                      @input="setCell(col.index, $event.target.value); autosize($event)"
                    ></textarea>
                    <button v-if="fileTreeOf(col) || renderedHtml(col)" type="button" class="html-toggle" @click.stop="htmlEdit[col.index] = false">{{ fileTreeOf(col) ? '渲染目录' : '渲染样式' }}</button>
                  </template>
                  <template v-else-if="widgetOf(col) === 'input'">
                    <input
                      class="text-input line-input"
                      :aria-label="col.label"
                      :value="current.cells[col.index]"
                      @input="setCell(col.index, $event.target.value)"
                    />
                    <button v-if="fileTreeOf(col) || renderedHtml(col)" type="button" class="html-toggle" @click.stop="htmlEdit[col.index] = false">{{ fileTreeOf(col) ? '渲染目录' : '渲染样式' }}</button>
                  </template>
                  <div v-else>
                    <p v-if="!col.options.length" class="empty-options">还没有可选项，可以直接输入。</p>
                    <div :class="['chips', { scroll: col.options.length > 12 }]" :role="col.mode === 'multi' ? 'group' : 'radiogroup'" :aria-label="col.label">
                      <button
                        v-for="option in chipOptions(col)"
                        :key="option.value"
                        type="button"
                        class="chip"
                        :class="{ on: isChosen(col, option) }"
                        :role="col.mode === 'multi' ? 'checkbox' : 'radio'"
                        :aria-checked="isChosen(col, option)"
                        @click="col.mode === 'multi' ? chooseMulti(col, option.value) : chooseSingle(col, option)"
                      >
                        {{ option.label }}
                      </button>
                      <input
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

        <div class="pane-foot detail-foot">
          <p v-if="bookType === 'xls'" class="hint">XLS 写回会重建工作簿，复杂格式可能变化。</p>
          <p v-else-if="fileHandle" class="hint">保存将覆盖源文件中改过的单元格。</p>
          <p v-else class="hint">当前文件没有写入句柄，保存时下载新文件。</p>
        </div>
      </main>
      </div>
    </div>

    <div v-if="toastText" class="toast" role="status">{{ toastText }}</div>
  </div>
</template>
