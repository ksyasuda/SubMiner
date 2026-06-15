import { createContext, useContext, useSyncExternalStore, useCallback } from 'react';
import { i18n } from './index';
import type { SupportedLanguage } from './types';

interface I18nContextValue {
  t: (key: string, params?: Record<string, string | number>) => string;
  language: SupportedLanguage;
  setLanguage: (lang: SupportedLanguage) => void;
}

const I18nContext = createContext<I18nContextValue | null>(null);

function subscribe(callback: () => void): () => void {
  i18n.onChange = callback;
  return () => {
    i18n.onChange = undefined;
  };
}

function getSnapshot(): SupportedLanguage {
  return i18n.getLanguage();
}

export function I18nProvider({
  children,
  initialLanguage,
}: {
  children: React.ReactNode;
  initialLanguage?: SupportedLanguage;
}) {
  if (initialLanguage) {
    i18n.setLanguage(initialLanguage);
  }

  const language = useSyncExternalStore(subscribe, getSnapshot, getSnapshot);

  const setLanguage = useCallback((lang: SupportedLanguage) => {
    i18n.setLanguage(lang);
  }, []);

  const t = useCallback(
    (key: string, params?: Record<string, string | number>) => i18n.t(key, params),
    [language],
  );

  return (
    <I18nContext.Provider value={{ t, language, setLanguage }}>
      {children}
    </I18nContext.Provider>
  );
}

export function useTranslation() {
  const context = useContext(I18nContext);
  if (!context) {
    throw new Error('useTranslation must be used within an I18nProvider');
  }
  return context;
}
