import type { PerAnimeDataPoint } from './StackedTrendChart';

const HIDDEN_TITLES_KEY = 'subminer-stats-trends-hidden-titles';
const MAX_TITLES_KEY = 'subminer-stats-trends-max-titles';

export const MAX_TITLES_OPTIONS = [3, 5, 7, 10] as const;

function getStorage(): Storage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

export function loadHiddenTitles(): Set<string> {
  try {
    const raw = getStorage()?.getItem(HIDDEN_TITLES_KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(parsed)) return new Set();
    return new Set(parsed.filter((value): value is string => typeof value === 'string'));
  } catch {
    return new Set();
  }
}

export function saveHiddenTitles(hidden: ReadonlySet<string>): void {
  try {
    getStorage()?.setItem(HIDDEN_TITLES_KEY, JSON.stringify([...hidden]));
  } catch {
    // Storage can be blocked in private/restricted contexts; keep the in-memory choice.
  }
}

export function loadMaxTitles(): number | null {
  try {
    const raw = getStorage()?.getItem(MAX_TITLES_KEY);
    if (raw === null || raw === undefined) return null;
    const value = Number(raw);
    return (MAX_TITLES_OPTIONS as readonly number[]).includes(value) ? value : null;
  } catch {
    return null;
  }
}

export function saveMaxTitles(value: number | null): void {
  try {
    const storage = getStorage();
    if (!storage) return;
    if (value === null) {
      storage.removeItem(MAX_TITLES_KEY);
    } else {
      storage.setItem(MAX_TITLES_KEY, String(value));
    }
  } catch {
    // Storage can be blocked in private/restricted contexts; keep the in-memory choice.
  }
}

export function buildAnimeVisibilityOptions(datasets: PerAnimeDataPoint[][]): string[] {
  const totals = new Map<string, number>();
  for (const dataset of datasets) {
    for (const point of dataset) {
      totals.set(point.animeTitle, (totals.get(point.animeTitle) ?? 0) + point.value);
    }
  }

  return [...totals.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .map(([title]) => title);
}

export function filterHiddenAnimeData(
  data: PerAnimeDataPoint[],
  hiddenAnime: ReadonlySet<string>,
): PerAnimeDataPoint[] {
  if (hiddenAnime.size === 0) {
    return data;
  }
  return data.filter((point) => !hiddenAnime.has(point.animeTitle));
}

export function pruneHiddenAnime(
  hiddenAnime: ReadonlySet<string>,
  availableAnime: readonly string[],
): Set<string> {
  const availableSet = new Set(availableAnime);
  return new Set([...hiddenAnime].filter((title) => availableSet.has(title)));
}
