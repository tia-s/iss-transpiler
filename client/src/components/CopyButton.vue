<template>
  <button
    class="copy-btn"
    :class="{ 'copy-btn--copied': copied }"
    :disabled="!code"
    @click="handleCopy"
  >
    <svg v-if="!copied" width="14" height="14" viewBox="0 0 14 14" fill="none">
      <rect x="4" y="4" width="8" height="8" rx="1" stroke="currentColor" stroke-width="1.5"/>
      <path d="M2 10V2h8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
    </svg>
    <svg v-else width="14" height="14" viewBox="0 0 14 14" fill="none">
      <path d="M2 7l4 4 6-7" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>
  </button>
</template>

<script setup lang="ts">
import { ref } from 'vue'

const props = defineProps<{
  code: string
}>()

const copied = ref(false)

async function handleCopy() {
  if (!props.code) return
  await navigator.clipboard.writeText(props.code)
  copied.value = true
  setTimeout(() => { copied.value = false }, 2000)
}
</script>

<style lang="scss" scoped>
@use '../styles/components/copy-button';
</style>