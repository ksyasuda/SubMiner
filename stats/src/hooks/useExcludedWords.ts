import { useCallback, useEffect, useSyncExternalStore } from 'react';
import { apiClient } from '../lib/api-client';
import type { StatsExcludedWord } from '../types/stats';

export type ExcludedWord = StatsExcludedWord;

const STORAGE_KEY = 'subminer-excluded-words';

type ExclusionCandidate = { headword: string; word: string; reading: string };

function normalizedTokenText(value: string): string {
  return value.trim();
}

export function getExcludedWordTokenKey(w: ExclusionCandidate): string {
  return (
    normalizedTokenText(w.headword) || normalizedTokenText(w.word) || normalizedTokenText(w.reading)
  );
}

function getExcludedWordAliasKeys(w: ExclusionCandidate): string[] {
  const aliases = [normalizedTokenText(w.headword), normalizedTokenText(w.word)].filter(Boolean);
  const unique = new Set(aliases);
  if (unique.size === 0) unique.add(getExcludedWordTokenKey(w));
  return [...unique];
}

export function dedupeExcludedWords(words: ExcludedWord[]): ExcludedWord[] {
  const byToken = new Map<string, ExcludedWord>();
  for (const word of words) {
    const key = getExcludedWordTokenKey(word);
    if (!byToken.has(key)) byToken.set(key, word);
  }
  return [...byToken.values()];
}

export function isExcludedWord(excluded: ExcludedWord[], w: ExclusionCandidate): boolean {
  const excludedKeys = new Set(excluded.flatMap(getExcludedWordAliasKeys));
  return getExcludedWordAliasKeys(w).some((key) => excludedKeys.has(key));
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
    return dedupeExcludedWords(
      parsed.filter(
        (row): row is ExcludedWord =>
          row !== null &&
          typeof row === 'object' &&
          typeof (row as ExcludedWord).headword === 'string' &&
          typeof (row as ExcludedWord).word === 'string' &&
          typeof (row as ExcludedWord).reading === 'string',
      ),
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
  cachedKeys = new Set(load().flatMap(getExcludedWordAliasKeys));
  return cachedKeys;
}

function applyWords(words: ExcludedWord[]): void {
  const normalized = dedupeExcludedWords(words);
  cached = normalized;
  cachedKeys = new Set(normalized.flatMap(getExcludedWordAliasKeys));
  writeLocalStorage(normalized);
  for (const fn of listeners) fn();
}

export function getExcludedWordsSnapshot(): ExcludedWord[] {
  return load();
}

export async function setExcludedWords(words: ExcludedWord[]): Promise<void> {
  const previousWords = [...load()];
  const previousRevision = revision;
  const writeRevision = previousRevision + 1;
  const normalized = dedupeExcludedWords(words);
  revision = writeRevision;
  applyWords(normalized);
  try {
    await apiClient.setExcludedWords(normalized);
  } catch (error) {
    if (revision === writeRevision) {
      revision = previousRevision;
      applyWords(previousWords);
    }
    console.error('Failed to persist excluded words to stats database', error);
    throw error;
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
      initialized = null;
      console.error('Failed to load excluded words from stats database', error);
      return;
    }

    if (revision !== startRevision) {
      initialized = null;
      return;
    }
    if (dbWords.length > 0) {
      applyWords(dbWords);
      return;
    }
    if (localWords.length > 0) {
      try {
        await setExcludedWords(localWords);
      } catch {
        initialized = null;
      }
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
    (w: ExclusionCandidate) => getExcludedWordAliasKeys(w).some((key) => getKeySet().has(key)),
    [excluded],
  );

  const toggleExclusion = useCallback((w: ExcludedWord) => {
    const current = load();
    const candidateKeys = new Set(getExcludedWordAliasKeys(w));
    const existing = current.find((e) =>
      getExcludedWordAliasKeys(e).some((key) => candidateKeys.has(key)),
    );
    if (existing) {
      const key = getExcludedWordTokenKey(existing);
      void setExcludedWords(current.filter((e) => getExcludedWordTokenKey(e) !== key));
    } else {
      void setExcludedWords([...current, w]);
    }
  }, []);

  const removeExclusion = useCallback((w: ExcludedWord) => {
    const key = getExcludedWordTokenKey(w);
    void setExcludedWords(load().filter((e) => getExcludedWordTokenKey(e) !== key));
  }, []);

  const clearAll = useCallback(() => {
    void setExcludedWords([]);
  }, []);

  return { excluded, isExcluded, toggleExclusion, removeExclusion, clearAll };
}
