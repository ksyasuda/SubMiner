import type { Hono } from 'hono';
import { statsJson } from '../../../types/stats-http-contract.js';
import type { ImmersionTrackerService } from '../immersion-tracker-service.js';
import {
  enrichSessionsWithKnownWordMetrics,
  loadKnownWordsSet,
  parseEventTypesQuery,
  parseIntQuery,
  parseTrendFillEmpty,
  parseTrendGroupBy,
  parseTrendRange,
} from './route-support.js';

export function registerStatsAnalyticsRoutes(
  app: Hono,
  tracker: ImmersionTrackerService,
  options?: { knownWordCachePath?: string },
): void {
  app.get('/api/stats/overview', async (c) => {
    const [rawSessions, rollups, hints] = await Promise.all([
      tracker.getSessionSummaries(5),
      tracker.getDailyRollups(14),
      tracker.getQueryHints(),
    ]);
    const sessions = await enrichSessionsWithKnownWordMetrics(
      tracker,
      rawSessions,
      options?.knownWordCachePath,
    );
    return c.json(statsJson('overview', { sessions, rollups, hints }));
  });

  app.get('/api/stats/daily-rollups', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 60, 500);
    const rollups = await tracker.getDailyRollups(limit);
    return c.json(statsJson('dailyRollups', rollups));
  });

  app.get('/api/stats/monthly-rollups', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 24, 120);
    const rollups = await tracker.getMonthlyRollups(limit);
    return c.json(statsJson('monthlyRollups', rollups));
  });

  app.get('/api/stats/streak-calendar', async (c) => {
    const days = parseIntQuery(c.req.query('days'), 90, 365);
    return c.json(statsJson('streakCalendar', await tracker.getStreakCalendar(days)));
  });

  app.get('/api/stats/trends/episodes-per-day', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 90, 365);
    return c.json(statsJson('episodesPerDay', await tracker.getEpisodesPerDay(limit)));
  });

  app.get('/api/stats/trends/new-anime-per-day', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 90, 365);
    return c.json(statsJson('newAnimePerDay', await tracker.getNewAnimePerDay(limit)));
  });

  app.get('/api/stats/trends/watch-time-per-anime', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 90, 365);
    return c.json(statsJson('watchTimePerAnime', await tracker.getWatchTimePerAnime(limit)));
  });

  app.get('/api/stats/trends/dashboard', async (c) => {
    const range = parseTrendRange(c.req.query('range'));
    const groupBy = parseTrendGroupBy(c.req.query('groupBy'));
    const fillEmpty = parseTrendFillEmpty(c.req.query('fillEmpty'));
    return c.json(
      statsJson('trendsDashboard', await tracker.getTrendsDashboard(range, groupBy, fillEmpty)),
    );
  });

  app.get('/api/stats/sessions', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 50, 500);
    const rawSessions = await tracker.getSessionSummaries(limit);
    const sessions = await enrichSessionsWithKnownWordMetrics(
      tracker,
      rawSessions,
      options?.knownWordCachePath,
    );
    return c.json(statsJson('sessions', sessions));
  });

  app.get('/api/stats/sessions/:id/timeline', async (c) => {
    const id = parseIntQuery(c.req.param('id'), 0);
    if (id <= 0) return c.json(statsJson('sessionTimeline', []), 400);
    const rawLimit = c.req.query('limit');
    const limit = rawLimit === undefined ? undefined : parseIntQuery(rawLimit, 200, 1000);
    const timeline = await tracker.getSessionTimeline(id, limit);
    return c.json(statsJson('sessionTimeline', timeline));
  });

  app.get('/api/stats/sessions/:id/events', async (c) => {
    const id = parseIntQuery(c.req.param('id'), 0);
    if (id <= 0) return c.json(statsJson('sessionEvents', []), 400);
    const limit = parseIntQuery(c.req.query('limit'), 500, 1000);
    const eventTypes = parseEventTypesQuery(c.req.query('types'));
    const events = await tracker.getSessionEvents(id, limit, eventTypes);
    return c.json(statsJson('sessionEvents', events));
  });

  app.get('/api/stats/sessions/:id/known-words-timeline', async (c) => {
    const id = parseIntQuery(c.req.param('id'), 0);
    if (id <= 0) return c.json(statsJson('sessionKnownWordsTimeline', []), 400);

    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath) ?? new Set<string>();

    // Get per-line word occurrences for the session.
    const wordsByLine = await tracker.getSessionWordsByLine(id);

    // Build cumulative filtered occurrence counts per recorded line index.
    // The stats UI uses line-count progress to align this series with the session
    // timeline, so preserve the stored line position rather than compressing gaps.
    const totalLineGroups = new Map<number, number>();
    const knownLineGroups = new Map<number, number>();
    for (const row of wordsByLine) {
      totalLineGroups.set(
        row.lineIndex,
        (totalLineGroups.get(row.lineIndex) ?? 0) + row.occurrenceCount,
      );
      if (knownWordsSet.has(row.headword)) {
        knownLineGroups.set(
          row.lineIndex,
          (knownLineGroups.get(row.lineIndex) ?? 0) + row.occurrenceCount,
        );
      }
    }

    const maxLineIndex = Math.max(...totalLineGroups.keys(), ...knownLineGroups.keys(), -1);
    let knownWordsSeen = 0;
    let totalWordsSeen = 0;
    const knownByLinesSeen: Array<{
      linesSeen: number;
      knownWordsSeen: number;
      totalWordsSeen: number;
    }> = [];

    for (let lineIdx = 0; lineIdx <= maxLineIndex; lineIdx += 1) {
      knownWordsSeen += knownLineGroups.get(lineIdx) ?? 0;
      totalWordsSeen += totalLineGroups.get(lineIdx) ?? 0;
      knownByLinesSeen.push({
        linesSeen: lineIdx,
        knownWordsSeen,
        totalWordsSeen,
      });
    }

    return c.json(statsJson('sessionKnownWordsTimeline', knownByLinesSeen));
  });
}
