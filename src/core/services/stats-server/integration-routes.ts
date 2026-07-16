import type { Hono } from 'hono';
import type { AnkiConnectConfig } from '../../../types.js';
import {
  statsJson,
  type StatsAnilistSearchResult,
  type StatsAnkiBrowseResponse,
} from '../../../types/stats-http-contract.js';
import type { AnilistRateLimiter } from '../anilist/rate-limiter.js';
import { registerStatsCoverRoutes } from '../stats-cover-routes.js';
import type { ImmersionTrackerService } from '../immersion-tracker-service.js';
import {
  buildAnkiNotePreview,
  countKnownWords,
  enrichSessionsWithKnownWordMetrics,
  loadKnownWordsSet,
  parseIntQuery,
} from './route-support.js';

const ANKI_CONNECT_FETCH_TIMEOUT_MS = 3_000;
const ANILIST_FETCH_TIMEOUT_MS = 3_000;

export function registerStatsIntegrationRoutes(
  app: Hono,
  tracker: ImmersionTrackerService,
  options?: {
    knownWordCachePath?: string;
    ankiConnectConfig?: AnkiConnectConfig;
    getAnkiConnectConfig?: () => AnkiConnectConfig | undefined;
    anilistRateLimiter?: AnilistRateLimiter;
    resolveAnkiNoteId?: (noteId: number) => number;
  },
): void {
  const getAnkiConnectConfig = (): AnkiConnectConfig | undefined =>
    options?.getAnkiConnectConfig?.() ?? options?.ankiConnectConfig;
  app.get('/api/stats/anilist/search', async (c) => {
    const query = (c.req.query('q') ?? '').trim();
    if (!query) return c.json(statsJson('anilistSearch', []));
    try {
      await options?.anilistRateLimiter?.acquire();
      const res = await fetch('https://graphql.anilist.co', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: AbortSignal.timeout(ANILIST_FETCH_TIMEOUT_MS),
        body: JSON.stringify({
          query: `query ($search: String!) {
          Page(perPage: 10) {
            media(search: $search, type: ANIME) {
              id
              episodes
              season
              seasonYear
              description(asHtml: false)
              coverImage { large medium }
              title { romaji english native }
            }
          }
        }`,
          variables: { search: query },
        }),
      });
      options?.anilistRateLimiter?.recordResponse(res.headers);
      if (res.status === 429) {
        return c.json(statsJson('anilistSearch', []));
      }
      const json = (await res.json()) as {
        data?: { Page?: { media?: StatsAnilistSearchResult[] } };
      };
      return c.json(statsJson('anilistSearch', json.data?.Page?.media ?? []));
    } catch {
      return c.json(statsJson('anilistSearch', []));
    }
  });

  app.get('/api/stats/known-words', (c) => {
    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath);
    if (!knownWordsSet) return c.json(statsJson('knownWords', []));
    return c.json(statsJson('knownWords', [...knownWordsSet]));
  });

  app.get('/api/stats/known-words-summary', async (c) => {
    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath);
    if (!knownWordsSet) {
      return c.json(statsJson('knownWordsSummary', { totalUniqueWords: 0, knownWordCount: 0 }));
    }
    const headwords = await tracker.getAllDistinctHeadwords();
    return c.json(statsJson('knownWordsSummary', countKnownWords(headwords, knownWordsSet)));
  });

  app.get('/api/stats/anime/:animeId/known-words-summary', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    if (animeId <= 0) {
      return c.json(
        statsJson('animeKnownWordsSummary', { totalUniqueWords: 0, knownWordCount: 0 }),
        400,
      );
    }
    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath);
    if (!knownWordsSet) {
      return c.json(
        statsJson('animeKnownWordsSummary', { totalUniqueWords: 0, knownWordCount: 0 }),
      );
    }
    const headwords = await tracker.getAnimeDistinctHeadwords(animeId);
    return c.json(statsJson('animeKnownWordsSummary', countKnownWords(headwords, knownWordsSet)));
  });

  app.get('/api/stats/media/:videoId/known-words-summary', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) {
      return c.json(
        statsJson('mediaKnownWordsSummary', { totalUniqueWords: 0, knownWordCount: 0 }),
        400,
      );
    }
    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath);
    if (!knownWordsSet) {
      return c.json(
        statsJson('mediaKnownWordsSummary', { totalUniqueWords: 0, knownWordCount: 0 }),
      );
    }
    const headwords = await tracker.getMediaDistinctHeadwords(videoId);
    return c.json(statsJson('mediaKnownWordsSummary', countKnownWords(headwords, knownWordsSet)));
  });

  app.patch('/api/stats/anime/:animeId/anilist', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    if (animeId <= 0) return c.body(null, 400);
    const body = await c.req.json().catch(() => null);
    if (
      typeof body?.anilistId !== 'number' ||
      !Number.isInteger(body.anilistId) ||
      body.anilistId <= 0
    ) {
      return c.body(null, 400);
    }
    await tracker.reassignAnimeAnilist(animeId, body);
    return c.json(statsJson('reassignAnimeAnilist', { ok: true }));
  });

  registerStatsCoverRoutes(app, tracker);

  app.get('/api/stats/episode/:videoId/detail', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.body(null, 400);
    const rawSessions = await tracker.getEpisodeSessions(videoId);
    const words = await tracker.getEpisodeWords(videoId);
    const cardEvents = await tracker.getEpisodeCardEvents(videoId);
    const sessions = await enrichSessionsWithKnownWordMetrics(
      tracker,
      rawSessions,
      options?.knownWordCachePath,
    );
    return c.json(statsJson('episodeDetail', { sessions, words, cardEvents }));
  });

  app.post('/api/stats/anki/browse', async (c) => {
    const noteId = parseIntQuery(c.req.query('noteId'), 0);
    if (noteId <= 0) return c.body(null, 400);
    const ankiConfig = getAnkiConnectConfig();
    try {
      const response = await fetch(ankiConfig?.url ?? 'http://127.0.0.1:8765', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: AbortSignal.timeout(ANKI_CONNECT_FETCH_TIMEOUT_MS),
        body: JSON.stringify({
          action: 'guiBrowse',
          version: 6,
          params: { query: `nid:${noteId}` },
        }),
      });
      const result = (await response.json()) as StatsAnkiBrowseResponse;
      return c.json(statsJson('ankiBrowse', result));
    } catch {
      return c.json(statsJson('error', { error: 'Failed to reach AnkiConnect' }), 502);
    }
  });

  app.post('/api/stats/anki/notesInfo', async (c) => {
    const body = await c.req.json().catch(() => null);
    const noteIds: number[] = Array.isArray(body?.noteIds)
      ? body.noteIds.filter(
          (id: unknown): id is number => typeof id === 'number' && Number.isInteger(id) && id > 0,
        )
      : [];
    if (noteIds.length === 0) return c.json(statsJson('ankiNotesInfo', []));
    const resolvedNoteIds = Array.from(
      new Set(
        noteIds.map((noteId) => {
          const resolvedNoteId = options?.resolveAnkiNoteId?.(noteId);
          return Number.isInteger(resolvedNoteId) && (resolvedNoteId as number) > 0
            ? (resolvedNoteId as number)
            : noteId;
        }),
      ),
    );
    try {
      const ankiConfig = getAnkiConnectConfig();
      const response = await fetch(ankiConfig?.url ?? 'http://127.0.0.1:8765', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: AbortSignal.timeout(ANKI_CONNECT_FETCH_TIMEOUT_MS),
        body: JSON.stringify({
          action: 'notesInfo',
          version: 6,
          params: { notes: resolvedNoteIds },
        }),
      });
      const result = (await response.json()) as {
        result?: Array<{ noteId: number; fields: Record<string, { value: string }> }>;
      };
      return c.json(
        statsJson(
          'ankiNotesInfo',
          (result.result ?? []).map((note) => ({
            ...note,
            preview: buildAnkiNotePreview(note.fields, ankiConfig),
          })),
        ),
      );
    } catch {
      return c.json(statsJson('ankiNotesInfo', []), 502);
    }
  });
}
