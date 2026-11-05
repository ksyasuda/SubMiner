import type { TrendsDashboardData } from '../../types/stats';
import type { PerAnimeDataPoint } from './StackedTrendChart';

type AnimeSeries = TrendsDashboardData['animeDaily'];

/**
 * Turns sparse per-title daily values into running totals: one point per
 * title for every calendar day from its first active day to the last active
 * day of any title, so the stacked area charts carry each total forward.
 */
export function buildCumulativePerAnime(points: PerAnimeDataPoint[]): PerAnimeDataPoint[] {
  const byAnime = new Map<string, Map<number, number>>();
  let maxDay = Number.NEGATIVE_INFINITY;

  for (const point of points) {
    const dayMap = byAnime.get(point.animeTitle) ?? new Map<number, number>();
    dayMap.set(point.epochDay, (dayMap.get(point.epochDay) ?? 0) + point.value);
    byAnime.set(point.animeTitle, dayMap);
    maxDay = Math.max(maxDay, point.epochDay);
  }

  const result: PerAnimeDataPoint[] = [];
  for (const [animeTitle, dayMap] of byAnime) {
    const firstDay = Math.min(...dayMap.keys());
    let cumulative = 0;
    for (let epochDay = firstDay; epochDay <= maxDay; epochDay += 1) {
      cumulative += dayMap.get(epochDay) ?? 0;
      result.push({ epochDay, animeTitle, value: cumulative });
    }
  }
  return result;
}

export function buildAnimeCumulative(daily: AnimeSeries): AnimeSeries {
  return {
    watchTime: buildCumulativePerAnime(daily.watchTime),
    episodes: buildCumulativePerAnime(daily.episodes),
    cards: buildCumulativePerAnime(daily.cards),
    words: buildCumulativePerAnime(daily.words),
  };
}
