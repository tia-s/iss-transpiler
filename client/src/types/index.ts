export type PaneLanguage = 'ideascript' | 'output'

export interface TranslateRequest {
  code: string
}

export interface TranslateResponse {
  output: string 
  language: string // language based on the Translator class on backend
  framework: string // framework based on the Translator class on backend
}