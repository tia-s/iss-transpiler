import { ref } from 'vue'
import { translate } from '@/services/transpiler'
import type { TranslateResponse } from '@/types'

export function useTranspiler() {
  const isTranslating = ref(false)
  const isDone = ref(false)
  const error = ref<string | null>(null)
  const result = ref<TranslateResponse | null>(null)

  async function runTranslation(code: string) {
    isTranslating.value = true
    isDone.value = false
    error.value = null
    result.value = null

    try {
      result.value = await translate({ code })
      isDone.value = true
    } catch (e: unknown) {
      error.value = e instanceof Error ? e.message : 'Something went wrong'
    } finally {
      isTranslating.value = false
    }
  }

  return {
    isTranslating,
    isDone,
    error,
    result,
    runTranslation,
  }
}