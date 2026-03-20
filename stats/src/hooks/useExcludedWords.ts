import { useCallback, useSyncExternalStore } from 'react';

export interface ExcludedWord {
  headword: string;
  word: string;
  reading: string;
}

const STORAGE_KEY = 'subminer-excluded-words';

function toKey(w: ExcludedWord): string {
  return `${w.headword}\0${w.word}\0${w.reading}`;
}

let cached: ExcludedWord[] | null = null;
let cachedKeys: Set<string> | null = null;
const listeners = new Set<() => void>();

function load(): ExcludedWord[] {
  if (cached) return cached;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    cached = raw ? JSON.parse(raw) : [];
  } catch {
    cached = [];
  }
  return cached!;
}

function getKeySet(): Set<string> {
  if (cachedKeys) return cachedKeys;
  cachedKeys = new Set(load().map(toKey));
  return cachedKeys;
}

function persist(words: ExcludedWord[]) {
  cached = words;
  cachedKeys = new Set(words.map(toKey));
  localStorage.setItem(STORAGE_KEY, JSON.stringify(words));
  for (const fn of listeners) fn();
}

function getSnapshot(): ExcludedWord[] {
  return load();
}

function subscribe(fn: () => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

export function useExcludedWords() {
  const excluded = useSyncExternalStore(subscribe, getSnapshot);

  const isExcluded = useCallback(
    (w: { headword: string; word: string; reading: string }) => getKeySet().has(toKey(w)),
    [excluded],
  );

  const toggleExclusion = useCallback((w: ExcludedWord) => {
    const key = toKey(w);
    const current = load();
    if (getKeySet().has(key)) {
      persist(current.filter((e) => toKey(e) !== key));
    } else {
      persist([...current, w]);
    }
  }, []);

  const removeExclusion = useCallback((w: ExcludedWord) => {
    persist(load().filter((e) => toKey(e) !== toKey(w)));
  }, []);

  const clearAll = useCallback(() => persist([]), []);

  return { excluded, isExcluded, toggleExclusion, removeExclusion, clearAll };
}
