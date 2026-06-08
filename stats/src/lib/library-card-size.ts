export type LibraryCardSize = 'sm' | 'md' | 'lg';

export const DEFAULT_LIBRARY_CARD_SIZE: LibraryCardSize = 'md';
export const LIBRARY_CARD_SIZE_STORAGE_KEY = 'subminer.stats.library.cardSize';

export function getLibraryCardSizeStorage(
  source: { localStorage: Storage } | null | undefined,
): Storage | null {
  try {
    return source?.localStorage ?? null;
  } catch {
    return null;
  }
}

export function readLibraryCardSizePreference(
  storage: Storage | null | undefined,
): LibraryCardSize {
  try {
    const value = storage?.getItem(LIBRARY_CARD_SIZE_STORAGE_KEY);
    return value === 'sm' || value === 'md' || value === 'lg' ? value : DEFAULT_LIBRARY_CARD_SIZE;
  } catch {
    return DEFAULT_LIBRARY_CARD_SIZE;
  }
}

export function writeLibraryCardSizePreference(
  storage: Storage | null | undefined,
  size: LibraryCardSize,
): void {
  try {
    storage?.setItem(LIBRARY_CARD_SIZE_STORAGE_KEY, size);
  } catch {
    // Storage can be blocked in private/restricted contexts; keep the in-memory choice.
  }
}
