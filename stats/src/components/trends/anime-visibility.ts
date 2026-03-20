import type { PerAnimeDataPoint } from './StackedTrendChart';

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
