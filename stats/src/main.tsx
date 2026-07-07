import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import '@fontsource-variable/geist';
import '@fontsource-variable/geist-mono';
import { App } from './App';
import { I18nProvider } from './i18n';
import { i18n } from './i18n';
import './styles/globals.css';

// Initialize i18n: prefer lang query param, then navigator.language, then fallback to en
const params = new URLSearchParams(window.location.search);
const queryLang = params.get('lang');
if (queryLang === 'zh-CN' || queryLang === 'en') {
  i18n.setLanguage(queryLang);
} else {
  i18n.setLanguage(i18n.detectSystemLanguage());
}

const isOverlay = new URLSearchParams(window.location.search).has('overlay');
if (isOverlay) {
  document.body.classList.add('overlay-mode');
}

const root = document.getElementById('root');
if (root) {
  createRoot(root).render(
    <StrictMode>
      <I18nProvider>
        <App />
      </I18nProvider>
    </StrictMode>,
  );
}
