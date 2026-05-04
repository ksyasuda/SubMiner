import { useCallback, useEffect, useSyncExternalStore } from 'react';
import { apiClient } from '../lib/api-client';
import type { StatsExcludedWord } from '../types/stats';

export type ExcludedWord = StatsExcludedWord;

const STORAGE_KEY = 'subminer-excluded-words';

function toKey(w: ExcludedWord): string {
  return `${w.headword}\0${w.word}\0${w.reading}`;
}

let cached: ExcludedWord[] | null = null;
let cachedKeys: Set<string> | null = null;
let initialized: Promise<void> | null = null;
let revision = 0;
const listeners = new Set<() => void>();

function readLocalStorage(): ExcludedWord[] {
  if (typeof localStorage === 'undefined') return [];
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(
      (row): row is ExcludedWord =>
        row !== null &&
        typeof row === 'object' &&
        typeof (row as ExcludedWord).headword === 'string' &&
        typeof (row as ExcludedWord).word === 'string' &&
        typeof (row as ExcludedWord).reading === 'string',
    );
  } catch {
    return [];
  }
}

function writeLocalStorage(words: ExcludedWord[]): void {
  if (typeof localStorage === 'undefined') return;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(words));
}

function load(): ExcludedWord[] {
  if (cached) return cached;
  cached = readLocalStorage();
  return cached!;
}

function getKeySet(): Set<string> {
  if (cachedKeys) return cachedKeys;
  cachedKeys = new Set(load().map(toKey));
  return cachedKeys;
}

function applyWords(words: ExcludedWord[]): void {
  cached = words;
  cachedKeys = new Set(words.map(toKey));
  writeLocalStorage(words);
  for (const fn of listeners) fn();
}

export function getExcludedWordsSnapshot(): ExcludedWord[] {
  return load();
}

export async function setExcludedWords(words: ExcludedWord[]): Promise<void> {
  revision += 1;
  applyWords(words);
  try {
    await apiClient.setExcludedWords(words);
  } catch (error) {
    console.error('Failed to persist excluded words to stats database', error);
  }
}

export function initializeExcludedWordsStore(): Promise<void> {
  if (initialized) return initialized;
  const startRevision = revision;
  initialized = (async () => {
    const localWords = load();
    let dbWords: ExcludedWord[];
    try {
      dbWords = await apiClient.getExcludedWords();
    } catch (error) {
      console.error('Failed to load excluded words from stats database', error);
      return;
    }

    if (revision !== startRevision) return;
    if (dbWords.length > 0) {
      applyWords(dbWords);
      return;
    }
    if (localWords.length > 0) {
      await setExcludedWords(localWords);
      return;
    }
    applyWords([]);
  })();
  return initialized;
}

export function resetExcludedWordsStoreForTests(): void {
  cached = null;
  cachedKeys = null;
  initialized = null;
  revision = 0;
  listeners.clear();
}

function subscribe(fn: () => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

export function useExcludedWords() {
  const excluded = useSyncExternalStore(subscribe, getExcludedWordsSnapshot);

  useEffect(() => {
    void initializeExcludedWordsStore();
  }, []);

  const isExcluded = useCallback(
    (w: { headword: string; word: string; reading: string }) => getKeySet().has(toKey(w)),
    [excluded],
  );

  const toggleExclusion = useCallback((w: ExcludedWord) => {
    const key = toKey(w);
    const current = load();
    if (getKeySet().has(key)) {
      void setExcludedWords(current.filter((e) => toKey(e) !== key));
    } else {
      void setExcludedWords([...current, w]);
    }
  }, []);

  const removeExclusion = useCallback((w: ExcludedWord) => {
    void setExcludedWords(load().filter((e) => toKey(e) !== toKey(w)));
  }, []);

  const clearAll = useCallback(() => {
    void setExcludedWords([]);
  }, []);

  return { excluded, isExcluded, toggleExclusion, removeExclusion, clearAll };
}
