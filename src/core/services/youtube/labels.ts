import type { YoutubeTrackKind } from './kinds';

export type { YoutubeTrackKind };

export function normalizeYoutubeLangCode(value: string): string {
  return value.trim().toLowerCase().replace(/_/g, '-').replace(/[^a-z0-9-]+/g, '');
}

export function isJapaneseYoutubeLang(value: string): boolean {
  const normalized = normalizeYoutubeLangCode(value);
  return (
    normalized === 'ja' ||
    normalized === 'jp' ||
    normalized === 'jpn' ||
    normalized === 'japanese' ||
    normalized.startsWith('ja-') ||
    normalized.startsWith('jp-')
  );
}

export function isEnglishYoutubeLang(value: string): boolean {
  const normalized = normalizeYoutubeLangCode(value);
  return (
    normalized === 'en' ||
    normalized === 'eng' ||
    normalized === 'english' ||
    normalized === 'enus' ||
    normalized === 'en-us' ||
    normalized.startsWith('en-')
  );
}

export function formatYoutubeTrackLabel(input: {
  language: string;
  kind: YoutubeTrackKind;
  title?: string;
}): string {
  const language = input.language.trim() || 'unknown';
  const base = input.title?.trim() || language;
  return `${base} (${input.kind})`;
}
