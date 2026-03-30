<template>
  <button
    class="translate-btn"
    :class="{
      'translate-btn--loading': loading,
      'translate-btn--done': done
    }"
    :disabled="loading"
    @click="$emit('translate')"
  >
    <span class="translate-btn__label">{{ label }}</span>
  </button>
</template>

<script setup lang="ts">
import { computed } from 'vue'

const props = defineProps<{
  loading: boolean
  done: boolean
}>()

defineEmits<{
  translate: []
}>()

const label = computed(() => {
  if (props.loading) return 'Translating…'
  if (props.done) return 'Done'
  return 'Translate'
})
</script>

<style lang="scss" scoped>
@use '../styles/components/translate-button';
</style>