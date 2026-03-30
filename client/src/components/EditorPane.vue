<template>
  <div class="editor-pane" :class="[`editor-pane--${language}`, { 'editor-pane--focused': isFocused }]">
    <div class="editor-pane__header">
      <div class="editor-pane__header-left">
        <slot name="icon" />
        <span class="editor-pane__label">{{ label }}</span>
      </div>
      <div class="editor-pane__header-right">
        <slot name="actions" />
      </div>
    </div>
    <div class="editor-pane__body">
      <div class="editor-pane__line-numbers" ref="lineNumbersRef">
        <div
          v-for="n in lineCount"
          :key="n"
          class="editor-pane__line-number"
        >{{ n }}</div>
      </div>
      <div class="editor-pane__editor">
        <textarea
          class="editor-pane__textarea"
          ref="textareaRef"
          v-model="model"
          :placeholder="placeholder"
          :readonly="readonly"
          spellcheck="false"
          @focus="isFocused = true"
          @blur="isFocused = false"
          @scroll="syncScroll"
          @keydown="$emit('keydown', $event)"
        />
        <pre
          class="editor-pane__highlight"
          aria-hidden="true"
        ><code ref="codeRef" :class="highlightClass">{{ model }}</code></pre>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onMounted, nextTick } from 'vue'
import hljs from 'highlight.js'
import type { PaneLanguage } from '@/types'

const props = defineProps<{
  label: string
  language: PaneLanguage
  highlightLanguage?: string
  placeholder?: string
  readonly?: boolean
}>()

defineEmits<{
  keydown: [e: KeyboardEvent]
}>()

const model = defineModel<string>({ required: true })
const isFocused = ref(false)
const textareaRef = ref<HTMLTextAreaElement | null>(null)
const lineNumbersRef = ref<HTMLElement | null>(null)
const codeRef = ref<HTMLElement | null>(null)

const lineCount = computed(() => model.value ? model.value.split('\n').length : 1)

const highlightClass = computed(() =>
  props.language === 'ideascript'
    ? 'language-vbscript'
    : `language-${props.highlightLanguage ?? 'plaintext'}`
)

function syncScroll() {
  if (textareaRef.value && lineNumbersRef.value) {
    lineNumbersRef.value.scrollTop = textareaRef.value.scrollTop
  }
}

watch(model, async () => {
  await nextTick()
  if (codeRef.value) hljs.highlightElement(codeRef.value)
})

onMounted(() => {
  if (codeRef.value) hljs.highlightElement(codeRef.value)
})
</script>

<style lang="scss" scoped>
@use '../styles/components/editor-pane';
</style>