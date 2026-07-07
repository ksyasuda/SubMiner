import type { PerAnimeDataPoint } from './StackedTrendChart';

const HIDDEN_TITLES_KEY = 'subminer-stats-trends-hidden-titles';
const MAX_TITLES_KEY = 'subminer-stats-trends-max-titles';
const MAX_TITLES_MODE_KEY = 'subminer-stats-trends-max-titles-mode';

export const MAX_TITLES_OPTIONS = [3, 5, 7, 10] as const;

// How the per-chart limit picks which titles survive: 'recent' keeps the most
// recently active titles, 'total' keeps the highest cumulative totals. 'recent'
// is listed first so it heads the dropdown.
export type MaxTitlesMode = 'total' | 'recent';
export const MAX_TITLES_MODES: readonly MaxTitlesMode[] = ['recent', 'total'];

// First-run defaults: show the 10 most recently active titles per chart.
const DEFAULT_MAX_TITLES = 10;
const DEFAULT_MAX_TITLES_MODE: MaxTitlesMode = 'recent';
const ALL_TITLES_STORED_VALUE = 'all';

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

// Returns the persisted per-chart limit: a specific count, or null for "All".
// Absent/garbage storage falls back to the default count (never null), so a
// first run shows the default cap rather than every title. "All" is stored
// explicitly so it round-trips instead of collapsing back to the default.
export function loadMaxTitles(): number | null {
  try {
    const raw = getStorage()?.getItem(MAX_TITLES_KEY);
    if (raw === ALL_TITLES_STORED_VALUE) return null;
    const value = Number(raw);
    return (MAX_TITLES_OPTIONS as readonly number[]).includes(value) ? value : DEFAULT_MAX_TITLES;
  } catch {
    return DEFAULT_MAX_TITLES;
  }
}

export function saveMaxTitles(value: number | null): void {
  try {
    getStorage()?.setItem(MAX_TITLES_KEY, value === null ? ALL_TITLES_STORED_VALUE : String(value));
  } catch {
    // Storage can be blocked in private/restricted contexts; keep the in-memory choice.
  }
}

export function loadMaxTitlesMode(): MaxTitlesMode {
  try {
    const raw = getStorage()?.getItem(MAX_TITLES_MODE_KEY);
    if (raw === 'recent' || raw === 'total') return raw;
    return DEFAULT_MAX_TITLES_MODE;
  } catch {
    return DEFAULT_MAX_TITLES_MODE;
  }
}

export function saveMaxTitlesMode(mode: MaxTitlesMode): void {
  try {
    getStorage()?.setItem(MAX_TITLES_MODE_KEY, mode);
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
