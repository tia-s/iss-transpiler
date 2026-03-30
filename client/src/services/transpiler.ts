import axios from 'axios'
import type { TranslateRequest, TranslateResponse } from '@/types'

const client = axios.create({
  baseURL: import.meta.env.VITE_API_TARGET,
  timeout: 30_000,
})

export async function translate(payload: TranslateRequest): Promise<TranslateResponse> {
  const { data } = await client.post<TranslateResponse>('/translate', payload)
  return data
}