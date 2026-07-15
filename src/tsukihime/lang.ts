// Pure language-code helpers shared between the main-process Tsukihime client
// and the overlay renderer. Must stay free of node imports so the renderer
// bundle can use it.

const LANG_FILENAME_SUFFIXES: Record<string, string> = {
  eng: 'en',
  jpn: 'ja',
  ger: 'de',
  deu: 'de',
  spa: 'es',
  por: 'pt',
  fre: 'fr',
  fra: 'fr',
  ita: 'it',
  rus: 'ru',
  ara: 'ar',
  chi: 'zh',
  zho: 'zh',
  kor: 'ko',
};

const LANG_DISPLAY_NAMES: Record<string, string> = {
  en: 'English',
  ja: 'Japanese',
  de: 'German',
  es: 'Spanish',
  pt: 'Portuguese',
  fr: 'French',
  it: 'Italian',
  ru: 'Russian',
  ar: 'Arabic',
  zh: 'Chinese',
  ko: 'Korean',
};

export function normalizeTsukihimeLangCode(code: string): string {
  const base = code.trim().toLowerCase().split(/[-_]/)[0] ?? '';
  return LANG_FILENAME_SUFFIXES[base] ?? base;
}

export function tsukihimeLangToFilenameSuffix(lang: string | undefined): string {
  const normalized = normalizeTsukihimeLangCode(lang ?? '');
  if (!normalized || normalized === 'und') return 'en';
  return normalized;
}

export function tsukihimeTrackMatchesLanguages(
  trackLang: string,
  configuredLanguages: string[],
): boolean {
  const normalizedTrack = normalizeTsukihimeLangCode(trackLang);
  // Unlabeled tracks cannot be classified; keep them visible rather than
  // hiding them behind a tab that never shows them.
  if (!normalizedTrack || normalizedTrack === 'und') return true;
  return configuredLanguages.some(
    (candidate) => normalizeTsukihimeLangCode(candidate) === normalizedTrack,
  );
}

export function describeTsukihimeTabLanguages(configuredLanguages: string[]): string {
  const normalized = [
    ...new Set(
      configuredLanguages.map((candidate) => normalizeTsukihimeLangCode(candidate)).filter(Boolean),
    ),
  ];
  if (normalized.length === 0) return 'English';
  return normalized.map((code) => LANG_DISPLAY_NAMES[code] ?? code.toUpperCase()).join(' / ');
}
