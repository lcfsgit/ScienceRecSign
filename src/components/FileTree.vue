<script setup>
import { computed, ref } from 'vue'
import { formatSize } from '../lib/tree.js'

const props = defineProps({
  node: { type: Object, required: true }
})

const open = ref(true)
const isDir = computed(() => props.node.dir || props.node.type === 'folder')
const meta = computed(() => {
  if (isDir.value) {
    const count = props.node.children?.length || 0
    return count ? `${count}` : ''
  }
  return formatSize(props.node.size)
})

function onClick() {
  if (isDir.value) open.value = !open.value
}
</script>

<template>
  <div class="tree-node">
    <button
      type="button"
      class="tree-row"
      :class="{ dir: isDir, file: !isDir }"
      :aria-expanded="isDir ? open : undefined"
      :title="node.path || node.fileName"
      @click.stop="onClick"
    >
      <span class="tree-twist" aria-hidden="true">{{ isDir ? (open ? '▾' : '▸') : '·' }}</span>
      <span class="tree-name">{{ node.fileName }}</span>
      <span v-if="meta" class="tree-meta">{{ meta }}</span>
    </button>
    <div v-if="isDir && open && node.children?.length" class="tree-children" role="group">
      <FileTree v-for="(child, index) in node.children" :key="child.id || child.path || `${child.fileName}-${index}`" :node="child" />
    </div>
  </div>
</template>
