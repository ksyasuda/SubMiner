import type { SupportedLanguage, TranslationMap, UILanguage } from './types';
import en from './locales/en.json';
import zhCN from './locales/zh-CN.json';

const BUNDLED_LOCALES: Record<SupportedLanguage, TranslationMap> = {
  en: en as TranslationMap,
  'zh-CN': zhCN as TranslationMap,
};

const SYSTEM_LOCALE_TO_LANG: Record<string, SupportedLanguage> = {
  'zh-CN': 'zh-CN',
  'zh': 'zh-CN',
  'zh-Hans': 'zh-CN',
  'zh-Hans-CN': 'zh-CN',
  'zh-SG': 'zh-CN',
  'zh-Hant': 'zh-CN',
  'zh-HK': 'zh-CN',
  'zh-TW': 'zh-CN',
  'zh-MO': 'zh-CN',
};

function interpolate(template: string, params?: Record<string, string | number>): string {
  if (!params) return template;
  return template.replace(/\{\{(\w+)\}\}/g, (_match, key: string) => {
    const value = params[key];
    return value !== undefined ? String(value) : `{{${key}}}`;
  });
}

class I18n {
  private currentLanguage: SupportedLanguage = 'en';
  private translations: TranslationMap = {};
  onChange?: () => void;

  constructor() {
    this.loadTranslations(this.currentLanguage);
  }

  setLanguage(lang: SupportedLanguage): void {
    this.currentLanguage = lang;
    this.loadTranslations(lang);
    if (typeof document !== 'undefined') {
      document.documentElement.lang = lang;
    }
    this.onChange?.();
  }

  getLanguage(): SupportedLanguage {
    return this.currentLanguage;
  }

  t(key: string, params?: Record<string, string | number>, defaultValue?: string): string {
    const value = this.translations[key];
    if (value !== undefined) return interpolate(value, params);
    const fallback = BUNDLED_LOCALES['en'][key];
    if (fallback !== undefined) return interpolate(fallback, params);
    return defaultValue ?? key;
  }

  detectSystemLanguage(): SupportedLanguage {
    if (typeof navigator !== 'undefined' && navigator.language) {
      const mapped = SYSTEM_LOCALE_TO_LANG[navigator.language];
      if (mapped) return mapped;
      const base = navigator.language.split('-')[0];
      if (base) {
        const baseMapped = SYSTEM_LOCALE_TO_LANG[base];
        if (baseMapped) return baseMapped;
      }
    }
    if (typeof process !== 'undefined' && typeof process.env !== 'undefined') {
      const env = process.env['LANG'] || process.env['LC_ALL'] || process.env['LC_MESSAGES'];
      if (env) {
        const parts = env.split('.')[0];
        if (parts) {
          const mapped = SYSTEM_LOCALE_TO_LANG[parts];
          if (mapped) return mapped;
          const base = parts.split('_')[0];
          if (base) {
            const baseMapped = SYSTEM_LOCALE_TO_LANG[base];
            if (baseMapped) return baseMapped;
          }
        }
      }
    }
    return 'en';
  }

  resolveLanguage(uiLanguage: UILanguage): SupportedLanguage {
    if (uiLanguage === 'system') {
      return this.detectSystemLanguage();
    }
    if (uiLanguage === 'en' || uiLanguage === 'zh-CN') {
      return uiLanguage;
    }
    return 'en';
  }

  private loadTranslations(lang: SupportedLanguage): void {
    this.translations = BUNDLED_LOCALES[lang] || BUNDLED_LOCALES['en'];
  }
}

export const i18n = new I18n();

export function t(
  key: string,
  params?: Record<string, string | number>,
  defaultValue?: string,
): string {
  return i18n.t(key, params, defaultValue);
}

export function applyI18nToDOM(root: Document | HTMLElement = document): void {
  const elements = root.querySelectorAll('[data-i18n]');
  elements.forEach((el) => {
    const key = el.getAttribute('data-i18n');
    if (key) {
      const paramsAttr = el.getAttribute('data-i18n-params');
      let params: Record<string, string | number> | undefined;
      if (paramsAttr) {
        try {
          params = JSON.parse(paramsAttr);
        } catch {
          // ignore malformed params
        }
      }
      if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
        const placeholderKey = el.getAttribute('data-i18n-placeholder');
        if (placeholderKey) {
          el.placeholder = i18n.t(placeholderKey, params);
        }
        if (el.getAttribute('data-i18n') === 'value') {
          el.value = i18n.t(key, params);
        }
      } else {
        el.textContent = i18n.t(key, params);
      }
    }
  });

  const placeholders = root.querySelectorAll('[data-i18n-placeholder]:not([data-i18n])');
  placeholders.forEach((el) => {
    if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
      const key = el.getAttribute('data-i18n-placeholder');
      if (key) {
        el.placeholder = i18n.t(key);
      }
    }
  });

  const ariaElements = root.querySelectorAll('[data-i18n-aria]');
  ariaElements.forEach((el) => {
    const key = el.getAttribute('data-i18n-aria');
    if (key) {
      el.setAttribute('aria-label', i18n.t(key));
    }
  });

  const titles = root.querySelectorAll('[data-i18n-title]');
  titles.forEach((el) => {
    const key = el.getAttribute('data-i18n-title');
    if (key) {
      (el as HTMLElement).title = i18n.t(key);
    }
  });
}
