<template>
  <div id="app" ref="appRef" @keydown="handleKeyDown" tabindex="0">
    <TranspilerHeader />
    <div class="editor-layout">
      <EditorPane
        v-model="ideaCode"
        label="IDEASCRIPT"
        language="ideascript"
        placeholder="Paste your IDEAScript code here…"
        @keydown="handleKeyDown"
      />
      <EditorPane
        v-model="outputCode"
        :label="outputLabel"
        language="output"
        :highlightLanguage="result?.language"
        placeholder="Output will appear here…"
        :readonly="isTranslating"
        @keydown="handleKeyDown"
      >
        <template #actions>
          <CopyButton :code="outputCode" />
        </template>
      </EditorPane>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'
import EditorPane from '@/components/EditorPane.vue'
import CopyButton from '@/components/CopyButton.vue'
import TranspilerHeader from '@/components/TranspilerHeader.vue'
import { useTranspiler } from '@/composables/useTranspiler'

const appRef = ref<HTMLElement | null>(null)
const ideaCode = ref('')
const outputCode = ref('')

const { isTranslating, result } = useTranspiler()

const outputLabel = computed(() =>
  result.value
    ? `${result.value.language.toUpperCase()} · ${result.value.framework}`
    : 'OUTPUT'
)

function handleKeyDown(e: KeyboardEvent) {
  if (e.shiftKey && e.key === 'Enter') {
    e.preventDefault()
  }
}
</script>