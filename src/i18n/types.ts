export type SupportedLanguage = 'en' | 'zh-CN';

export type UILanguage = SupportedLanguage | 'system';

export type TranslationValue = string;

export type TranslationMap = Record<string, TranslationValue>;

export interface I18nConfig {
  defaultLanguage: SupportedLanguage;
  fallbackLanguage: SupportedLanguage;
}
