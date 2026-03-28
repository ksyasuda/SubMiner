import { Hono } from 'hono';
import type { ImmersionTrackerService } from './immersion-tracker-service.js';
import { basename, extname, resolve, sep } from 'node:path';
import { readFileSync, existsSync, statSync } from 'node:fs';
import { MediaGenerator } from '../../media-generator.js';
import { AnkiConnectClient } from '../../anki-connect.js';
import type { AnkiConnectConfig } from '../../types.js';
import {
  getConfiguredSentenceFieldName,
  getConfiguredTranslationFieldName,
  getConfiguredWordFieldName,
  getPreferredNoteFieldValue,
} from '../../anki-field-config.js';
import { resolveAnimatedImageLeadInSeconds } from '../../anki-integration/animated-image-sync.js';

type StatsServerNoteInfo = {
  noteId: number;
  fields: Record<string, { value: string }>;
};

function parseIntQuery(raw: string | undefined, fallback: number, maxLimit?: number): number {
  if (raw === undefined) return fallback;
  const n = Number(raw);
  if (!Number.isFinite(n) || n < 0) {
    return fallback;
  }
  const parsed = Math.floor(n);
  return maxLimit === undefined ? parsed : Math.min(parsed, maxLimit);
}

function parseTrendRange(raw: string | undefined): '7d' | '30d' | '90d' | 'all' {
  return raw === '7d' || raw === '30d' || raw === '90d' || raw === 'all' ? raw : '30d';
}

function parseTrendGroupBy(raw: string | undefined): 'day' | 'month' {
  return raw === 'month' ? 'month' : 'day';
}

function parseEventTypesQuery(raw: string | undefined): number[] | undefined {
  if (!raw) return undefined;
  const parsed = raw
    .split(',')
    .map((entry) => Number.parseInt(entry.trim(), 10))
    .filter((entry) => Number.isInteger(entry) && entry > 0);
  return parsed.length > 0 ? parsed : undefined;
}

function resolveStatsNoteFieldName(
  noteInfo: StatsServerNoteInfo,
  ...preferredNames: (string | undefined)[]
): string | null {
  for (const preferredName of preferredNames) {
    if (!preferredName) continue;
    const resolved = Object.keys(noteInfo.fields).find(
      (fieldName) => fieldName.toLowerCase() === preferredName.toLowerCase(),
    );
    if (resolved) return resolved;
  }
  return null;
}

/** Load known words cache from disk into a Set. Returns null if unavailable. */
function loadKnownWordsSet(cachePath: string | undefined): Set<string> | null {
  if (!cachePath || !existsSync(cachePath)) return null;
  try {
    const raw = JSON.parse(readFileSync(cachePath, 'utf-8')) as {
      version?: number;
      words?: string[];
    };
    if ((raw.version === 1 || raw.version === 2) && Array.isArray(raw.words)) {
      return new Set(raw.words);
    }
  } catch {
    /* ignore */
  }
  return null;
}

/** Count how many headwords in the given list are in the known words set. */
function countKnownWords(
  headwords: string[],
  knownWordsSet: Set<string>,
): { totalUniqueWords: number; knownWordCount: number } {
  let knownWordCount = 0;
  for (const hw of headwords) {
    if (knownWordsSet.has(hw)) knownWordCount++;
  }
  return { totalUniqueWords: headwords.length, knownWordCount };
}

function toKnownWordRate(knownWordsSeen: number, tokensSeen: number): number {
  if (!Number.isFinite(knownWordsSeen) || !Number.isFinite(tokensSeen) || tokensSeen <= 0) {
    return 0;
  }
  return Number(((knownWordsSeen / tokensSeen) * 100).toFixed(1));
}

async function enrichSessionsWithKnownWordMetrics(
  tracker: ImmersionTrackerService,
  sessions: Array<{
    sessionId: number;
    tokensSeen: number;
  }>,
  knownWordsCachePath?: string,
): Promise<
  Array<{
    sessionId: number;
    tokensSeen: number;
    knownWordsSeen: number;
    knownWordRate: number;
  }>
> {
  const knownWordsSet = loadKnownWordsSet(knownWordsCachePath);
  if (!knownWordsSet) {
    return sessions.map((session) => ({
      ...session,
      knownWordsSeen: 0,
      knownWordRate: 0,
    }));
  }

  const enriched = await Promise.all(
    sessions.map(async (session) => {
      let knownWordsSeen = 0;
      try {
        const wordsByLine = await tracker.getSessionWordsByLine(session.sessionId);
        for (const row of wordsByLine) {
          if (knownWordsSet.has(row.headword)) {
            knownWordsSeen += row.occurrenceCount;
          }
        }
      } catch {
        knownWordsSeen = 0;
      }

      return {
        ...session,
        knownWordsSeen,
        knownWordRate: toKnownWordRate(knownWordsSeen, session.tokensSeen),
      };
    }),
  );

  return enriched;
}

export interface StatsServerConfig {
  port: number;
  staticDir: string; // Path to stats/dist/
  tracker: ImmersionTrackerService;
  knownWordCachePath?: string;
  mpvSocketPath?: string;
  ankiConnectConfig?: AnkiConnectConfig;
  addYomitanNote?: (word: string) => Promise<number | null>;
  resolveAnkiNoteId?: (noteId: number) => number;
}

const STATS_STATIC_CONTENT_TYPES: Record<string, string> = {
  '.css': 'text/css; charset=utf-8',
  '.gif': 'image/gif',
  '.html': 'text/html; charset=utf-8',
  '.ico': 'image/x-icon',
  '.jpeg': 'image/jpeg',
  '.jpg': 'image/jpeg',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.txt': 'text/plain; charset=utf-8',
  '.webp': 'image/webp',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};
const ANKI_CONNECT_FETCH_TIMEOUT_MS = 3_000;

function buildAnkiNotePreview(
  fields: Record<string, { value: string }>,
  ankiConfig?: Pick<AnkiConnectConfig, 'fields'>,
): { word: string; sentence: string; translation: string } {
  return {
    word: getPreferredNoteFieldValue(fields, [getConfiguredWordFieldName(ankiConfig)]),
    sentence: getPreferredNoteFieldValue(fields, [getConfiguredSentenceFieldName(ankiConfig)]),
    translation: getPreferredNoteFieldValue(fields, [
      getConfiguredTranslationFieldName(ankiConfig),
    ]),
  };
}

function resolveStatsStaticPath(staticDir: string, requestPath: string): string | null {
  const normalizedPath = requestPath.replace(/^\/+/, '') || 'index.html';
  const decodedPath = decodeURIComponent(normalizedPath);
  const absoluteStaticDir = resolve(staticDir);
  const absolutePath = resolve(absoluteStaticDir, decodedPath);
  if (
    absolutePath !== absoluteStaticDir &&
    !absolutePath.startsWith(`${absoluteStaticDir}${sep}`)
  ) {
    return null;
  }
  if (!existsSync(absolutePath)) {
    return null;
  }
  const stats = statSync(absolutePath);
  if (!stats.isFile()) {
    return null;
  }
  return absolutePath;
}

function createStatsStaticResponse(staticDir: string, requestPath: string): Response | null {
  const absolutePath = resolveStatsStaticPath(staticDir, requestPath);
  if (!absolutePath) {
    return null;
  }

  const extension = extname(absolutePath).toLowerCase();
  const contentType = STATS_STATIC_CONTENT_TYPES[extension] ?? 'application/octet-stream';
  const body = readFileSync(absolutePath);
  return new Response(body, {
    headers: {
      'Content-Type': contentType,
      'Cache-Control': absolutePath.endsWith('index.html')
        ? 'no-cache'
        : 'public, max-age=31536000, immutable',
    },
  });
}

export function createStatsApp(
  tracker: ImmersionTrackerService,
  options?: {
    staticDir?: string;
    knownWordCachePath?: string;
    mpvSocketPath?: string;
    ankiConnectConfig?: AnkiConnectConfig;
    addYomitanNote?: (word: string) => Promise<number | null>;
    resolveAnkiNoteId?: (noteId: number) => number;
  },
) {
  const app = new Hono();

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
    return c.json({ sessions, rollups, hints });
  });

  app.get('/api/stats/daily-rollups', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 60, 500);
    const rollups = await tracker.getDailyRollups(limit);
    return c.json(rollups);
  });

  app.get('/api/stats/monthly-rollups', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 24, 120);
    const rollups = await tracker.getMonthlyRollups(limit);
    return c.json(rollups);
  });

  app.get('/api/stats/streak-calendar', async (c) => {
    const days = parseIntQuery(c.req.query('days'), 90, 365);
    return c.json(await tracker.getStreakCalendar(days));
  });

  app.get('/api/stats/trends/episodes-per-day', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 90, 365);
    return c.json(await tracker.getEpisodesPerDay(limit));
  });

  app.get('/api/stats/trends/new-anime-per-day', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 90, 365);
    return c.json(await tracker.getNewAnimePerDay(limit));
  });

  app.get('/api/stats/trends/watch-time-per-anime', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 90, 365);
    return c.json(await tracker.getWatchTimePerAnime(limit));
  });

  app.get('/api/stats/trends/dashboard', async (c) => {
    const range = parseTrendRange(c.req.query('range'));
    const groupBy = parseTrendGroupBy(c.req.query('groupBy'));
    return c.json(await tracker.getTrendsDashboard(range, groupBy));
  });

  app.get('/api/stats/sessions', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 50, 500);
    const rawSessions = await tracker.getSessionSummaries(limit);
    const sessions = await enrichSessionsWithKnownWordMetrics(
      tracker,
      rawSessions,
      options?.knownWordCachePath,
    );
    return c.json(sessions);
  });

  app.get('/api/stats/sessions/:id/timeline', async (c) => {
    const id = parseIntQuery(c.req.param('id'), 0);
    if (id <= 0) return c.json([], 400);
    const rawLimit = c.req.query('limit');
    const limit = rawLimit === undefined ? undefined : parseIntQuery(rawLimit, 200, 1000);
    const timeline = await tracker.getSessionTimeline(id, limit);
    return c.json(timeline);
  });

  app.get('/api/stats/sessions/:id/events', async (c) => {
    const id = parseIntQuery(c.req.param('id'), 0);
    if (id <= 0) return c.json([], 400);
    const limit = parseIntQuery(c.req.query('limit'), 500, 1000);
    const eventTypes = parseEventTypesQuery(c.req.query('types'));
    const events = await tracker.getSessionEvents(id, limit, eventTypes);
    return c.json(events);
  });

  app.get('/api/stats/sessions/:id/known-words-timeline', async (c) => {
    const id = parseIntQuery(c.req.param('id'), 0);
    if (id <= 0) return c.json([], 400);

    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath);
    if (!knownWordsSet) return c.json([]);

    // Get per-line word occurrences for the session.
    const wordsByLine = await tracker.getSessionWordsByLine(id);

    // Build cumulative known-word occurrence count per recorded line index.
    // The stats UI uses line-count progress to align this series with the session
    // timeline, so preserve the stored line position rather than compressing gaps.
    const lineGroups = new Map<number, number>();
    for (const row of wordsByLine) {
      if (!knownWordsSet.has(row.headword)) {
        continue;
      }
      lineGroups.set(row.lineIndex, (lineGroups.get(row.lineIndex) ?? 0) + row.occurrenceCount);
    }

    const sortedLineIndices = [...lineGroups.keys()].sort((a, b) => a - b);
    let knownWordsSeen = 0;
    const knownByLinesSeen: Array<{ linesSeen: number; knownWordsSeen: number }> = [];

    for (const lineIdx of sortedLineIndices) {
      knownWordsSeen += lineGroups.get(lineIdx)!;
      knownByLinesSeen.push({
        linesSeen: lineIdx,
        knownWordsSeen,
      });
    }

    return c.json(knownByLinesSeen);
  });

  app.get('/api/stats/vocabulary', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 100, 500);
    const excludePos = c.req.query('excludePos')?.split(',').filter(Boolean);
    const vocab = await tracker.getVocabularyStats(limit, excludePos);
    return c.json(vocab);
  });

  app.get('/api/stats/vocabulary/occurrences', async (c) => {
    const headword = (c.req.query('headword') ?? '').trim();
    const word = (c.req.query('word') ?? '').trim();
    const reading = (c.req.query('reading') ?? '').trim();
    if (!headword || !word) {
      return c.json([], 400);
    }
    const limit = parseIntQuery(c.req.query('limit'), 50, 500);
    const offset = parseIntQuery(c.req.query('offset'), 0, 10_000);
    const occurrences = await tracker.getWordOccurrences(headword, word, reading, limit, offset);
    return c.json(occurrences);
  });

  app.get('/api/stats/kanji', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 100, 500);
    const kanji = await tracker.getKanjiStats(limit);
    return c.json(kanji);
  });

  app.get('/api/stats/kanji/occurrences', async (c) => {
    const kanji = (c.req.query('kanji') ?? '').trim();
    if (!kanji) {
      return c.json([], 400);
    }
    const limit = parseIntQuery(c.req.query('limit'), 50, 500);
    const offset = parseIntQuery(c.req.query('offset'), 0, 10_000);
    const occurrences = await tracker.getKanjiOccurrences(kanji, limit, offset);
    return c.json(occurrences);
  });

  app.get('/api/stats/vocabulary/:wordId/detail', async (c) => {
    const wordId = parseIntQuery(c.req.param('wordId'), 0);
    if (wordId <= 0) return c.body(null, 400);
    const detail = await tracker.getWordDetail(wordId);
    if (!detail) return c.body(null, 404);
    const animeAppearances = await tracker.getWordAnimeAppearances(wordId);
    const similarWords = await tracker.getSimilarWords(wordId);
    return c.json({ detail, animeAppearances, similarWords });
  });

  app.get('/api/stats/kanji/:kanjiId/detail', async (c) => {
    const kanjiId = parseIntQuery(c.req.param('kanjiId'), 0);
    if (kanjiId <= 0) return c.body(null, 400);
    const detail = await tracker.getKanjiDetail(kanjiId);
    if (!detail) return c.body(null, 404);
    const animeAppearances = await tracker.getKanjiAnimeAppearances(kanjiId);
    const words = await tracker.getKanjiWords(kanjiId);
    return c.json({ detail, animeAppearances, words });
  });

  app.get('/api/stats/media', async (c) => {
    const library = await tracker.getMediaLibrary();
    return c.json(library);
  });

  app.get('/api/stats/media/:videoId', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.json(null, 400);
    const [detail, rawSessions, rollups] = await Promise.all([
      tracker.getMediaDetail(videoId),
      tracker.getMediaSessions(videoId, 100),
      tracker.getMediaDailyRollups(videoId, 90),
    ]);
    const sessions = await enrichSessionsWithKnownWordMetrics(
      tracker,
      rawSessions,
      options?.knownWordCachePath,
    );
    return c.json({ detail, sessions, rollups });
  });

  app.get('/api/stats/anime', async (c) => {
    const rows = await tracker.getAnimeLibrary();
    return c.json(rows);
  });

  app.get('/api/stats/anime/:animeId', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    if (animeId <= 0) return c.body(null, 400);
    const detail = await tracker.getAnimeDetail(animeId);
    if (!detail) return c.body(null, 404);
    const [episodes, anilistEntries] = await Promise.all([
      tracker.getAnimeEpisodes(animeId),
      tracker.getAnimeAnilistEntries(animeId),
    ]);
    return c.json({ detail, episodes, anilistEntries });
  });

  app.get('/api/stats/anime/:animeId/words', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    const limit = parseIntQuery(c.req.query('limit'), 50, 200);
    if (animeId <= 0) return c.body(null, 400);
    return c.json(await tracker.getAnimeWords(animeId, limit));
  });

  app.get('/api/stats/anime/:animeId/rollups', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    const limit = parseIntQuery(c.req.query('limit'), 90, 365);
    if (animeId <= 0) return c.body(null, 400);
    return c.json(await tracker.getAnimeDailyRollups(animeId, limit));
  });

  app.patch('/api/stats/media/:videoId/watched', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.body(null, 400);
    const body = await c.req.json().catch(() => null);
    const watched = typeof body?.watched === 'boolean' ? body.watched : true;
    await tracker.setVideoWatched(videoId, watched);
    return c.json({ ok: true });
  });

  app.delete('/api/stats/sessions', async (c) => {
    const body = await c.req.json().catch(() => null);
    const ids = Array.isArray(body?.sessionIds)
      ? body.sessionIds.filter((id: unknown) => typeof id === 'number' && id > 0)
      : [];
    if (ids.length === 0) return c.body(null, 400);
    await tracker.deleteSessions(ids);
    return c.json({ ok: true });
  });

  app.delete('/api/stats/sessions/:sessionId', async (c) => {
    const sessionId = parseIntQuery(c.req.param('sessionId'), 0);
    if (sessionId <= 0) return c.body(null, 400);
    await tracker.deleteSession(sessionId);
    return c.json({ ok: true });
  });

  app.delete('/api/stats/media/:videoId', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.body(null, 400);
    await tracker.deleteVideo(videoId);
    return c.json({ ok: true });
  });

  app.get('/api/stats/anilist/search', async (c) => {
    const query = (c.req.query('q') ?? '').trim();
    if (!query) return c.json([]);
    try {
      const res = await fetch('https://graphql.anilist.co', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
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
      const json = (await res.json()) as { data?: { Page?: { media?: unknown[] } } };
      return c.json(json.data?.Page?.media ?? []);
    } catch {
      return c.json([]);
    }
  });

  app.get('/api/stats/known-words', (c) => {
    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath);
    if (!knownWordsSet) return c.json([]);
    return c.json([...knownWordsSet]);
  });

  app.get('/api/stats/known-words-summary', async (c) => {
    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath);
    if (!knownWordsSet) return c.json({ totalUniqueWords: 0, knownWordCount: 0 });
    const headwords = await tracker.getAllDistinctHeadwords();
    return c.json(countKnownWords(headwords, knownWordsSet));
  });

  app.get('/api/stats/anime/:animeId/known-words-summary', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    if (animeId <= 0) return c.json({ totalUniqueWords: 0, knownWordCount: 0 }, 400);
    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath);
    if (!knownWordsSet) return c.json({ totalUniqueWords: 0, knownWordCount: 0 });
    const headwords = await tracker.getAnimeDistinctHeadwords(animeId);
    return c.json(countKnownWords(headwords, knownWordsSet));
  });

  app.get('/api/stats/media/:videoId/known-words-summary', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.json({ totalUniqueWords: 0, knownWordCount: 0 }, 400);
    const knownWordsSet = loadKnownWordsSet(options?.knownWordCachePath);
    if (!knownWordsSet) return c.json({ totalUniqueWords: 0, knownWordCount: 0 });
    const headwords = await tracker.getMediaDistinctHeadwords(videoId);
    return c.json(countKnownWords(headwords, knownWordsSet));
  });

  app.patch('/api/stats/anime/:animeId/anilist', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    if (animeId <= 0) return c.body(null, 400);
    const body = await c.req.json().catch(() => null);
    if (!body?.anilistId) return c.body(null, 400);
    await tracker.reassignAnimeAnilist(animeId, body);
    return c.json({ ok: true });
  });

  app.get('/api/stats/anime/:animeId/cover', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    if (animeId <= 0) return c.body(null, 404);
    const art = await tracker.getAnimeCoverArt(animeId);
    if (!art?.coverBlob) return c.body(null, 404);
    return new Response(new Uint8Array(art.coverBlob), {
      headers: {
        'Content-Type': 'image/jpeg',
        'Cache-Control': 'public, max-age=86400',
      },
    });
  });

  app.get('/api/stats/media/:videoId/cover', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.body(null, 404);
    let art = await tracker.getCoverArt(videoId);
    if (!art?.coverBlob) {
      await tracker.ensureCoverArt(videoId);
      art = await tracker.getCoverArt(videoId);
    }
    if (!art?.coverBlob) return c.body(null, 404);
    return new Response(new Uint8Array(art.coverBlob), {
      headers: {
        'Content-Type': 'image/jpeg',
        'Cache-Control': 'public, max-age=604800',
      },
    });
  });

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
    return c.json({ sessions, words, cardEvents });
  });

  app.post('/api/stats/anki/browse', async (c) => {
    const noteId = parseIntQuery(c.req.query('noteId'), 0);
    if (noteId <= 0) return c.body(null, 400);
    try {
      const response = await fetch('http://127.0.0.1:8765', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: AbortSignal.timeout(ANKI_CONNECT_FETCH_TIMEOUT_MS),
        body: JSON.stringify({
          action: 'guiBrowse',
          version: 6,
          params: { query: `nid:${noteId}` },
        }),
      });
      const result = await response.json();
      return c.json(result);
    } catch {
      return c.json({ error: 'Failed to reach AnkiConnect' }, 502);
    }
  });

  app.post('/api/stats/anki/notesInfo', async (c) => {
    const body = await c.req.json().catch(() => null);
    const noteIds: number[] = Array.isArray(body?.noteIds)
      ? body.noteIds.filter(
          (id: unknown): id is number => typeof id === 'number' && Number.isInteger(id) && id > 0,
        )
      : [];
    if (noteIds.length === 0) return c.json([]);
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
      const response = await fetch('http://127.0.0.1:8765', {
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
        (result.result ?? []).map((note) => ({
          ...note,
          preview: buildAnkiNotePreview(note.fields, options?.ankiConnectConfig),
        })),
      );
    } catch {
      return c.json([], 502);
    }
  });

  app.post('/api/stats/mine-card', async (c) => {
    const body = await c.req.json().catch(() => null);
    const sourcePath = typeof body?.sourcePath === 'string' ? body.sourcePath.trim() : '';
    const startMs = typeof body?.startMs === 'number' ? body.startMs : NaN;
    const endMs = typeof body?.endMs === 'number' ? body.endMs : NaN;
    const sentence = typeof body?.sentence === 'string' ? body.sentence.trim() : '';
    const word = typeof body?.word === 'string' ? body.word.trim() : '';
    const secondaryText = typeof body?.secondaryText === 'string' ? body.secondaryText.trim() : '';
    const videoTitle = typeof body?.videoTitle === 'string' ? body.videoTitle.trim() : '';
    const rawMode = c.req.query('mode');
    const mode = rawMode === 'audio' ? 'audio' : rawMode === 'word' ? 'word' : 'sentence';

    if (!sourcePath || !sentence || !Number.isFinite(startMs) || !Number.isFinite(endMs)) {
      return c.json({ error: 'sourcePath, sentence, startMs, and endMs are required' }, 400);
    }

    if (!existsSync(sourcePath)) {
      return c.json({ error: 'File not found' }, 404);
    }

    const ankiConfig = options?.ankiConnectConfig;
    if (!ankiConfig) {
      return c.json({ error: 'AnkiConnect is not configured' }, 500);
    }

    const client = new AnkiConnectClient(ankiConfig.url ?? 'http://127.0.0.1:8765');
    const mediaGen = new MediaGenerator();

    const audioPadding = ankiConfig.media?.audioPadding ?? 0.5;
    const maxMediaDuration = ankiConfig.media?.maxMediaDuration ?? 30;

    const startSec = startMs / 1000;
    const endSec = endMs / 1000;
    const rawDuration = endSec - startSec;
    const clampedEndSec = rawDuration > maxMediaDuration ? startSec + maxMediaDuration : endSec;

    const highlightedSentence = word
      ? sentence.replace(
          new RegExp(word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'),
          `<b>${word}</b>`,
        )
      : sentence;

    const generateAudio = ankiConfig.media?.generateAudio !== false;
    const generateImage = ankiConfig.media?.generateImage !== false && mode !== 'audio';
    const imageType = ankiConfig.media?.imageType ?? 'static';
    const syncAnimatedImageToWordAudio =
      imageType === 'avif' && ankiConfig.media?.syncAnimatedImageToWordAudio !== false;

    const audioPromise = generateAudio
      ? mediaGen.generateAudio(sourcePath, startSec, clampedEndSec, audioPadding)
      : Promise.resolve(null);

    const createImagePromise = (animatedLeadInSeconds = 0): Promise<Buffer | null> => {
      if (!generateImage) {
        return Promise.resolve(null);
      }

      if (imageType === 'avif') {
        return mediaGen.generateAnimatedImage(sourcePath, startSec, clampedEndSec, audioPadding, {
          fps: ankiConfig.media?.animatedFps ?? 10,
          maxWidth: ankiConfig.media?.animatedMaxWidth ?? 640,
          maxHeight: ankiConfig.media?.animatedMaxHeight,
          crf: ankiConfig.media?.animatedCrf ?? 35,
          leadingStillDuration: animatedLeadInSeconds,
        });
      }

      const midpointSec = (startSec + clampedEndSec) / 2;
      return mediaGen.generateScreenshot(sourcePath, midpointSec, {
        format: ankiConfig.media?.imageFormat ?? 'jpg',
        quality: ankiConfig.media?.imageQuality ?? 92,
        maxWidth: ankiConfig.media?.imageMaxWidth,
        maxHeight: ankiConfig.media?.imageMaxHeight,
      });
    };

    const imagePromise =
      mode === 'word' && syncAnimatedImageToWordAudio
        ? Promise.resolve<Buffer | null>(null)
        : createImagePromise();

    const errors: string[] = [];
    let noteId: number;

    if (mode === 'word') {
      if (!options?.addYomitanNote) {
        return c.json({ error: 'Yomitan bridge not available' }, 500);
      }

      const [yomitanResult, audioResult, imageResult] = await Promise.allSettled([
        options.addYomitanNote(word),
        audioPromise,
        imagePromise,
      ]);

      if (yomitanResult.status === 'rejected' || !yomitanResult.value) {
        return c.json(
          {
            error: `Yomitan failed to create note: ${yomitanResult.status === 'rejected' ? (yomitanResult.reason as Error).message : 'no result'}`,
          },
          502,
        );
      }

      noteId = yomitanResult.value;
      const audioBuffer = audioResult.status === 'fulfilled' ? audioResult.value : null;
      if (audioResult.status === 'rejected')
        errors.push(`audio: ${(audioResult.reason as Error).message}`);
      if (imageResult.status === 'rejected')
        errors.push(`image: ${(imageResult.reason as Error).message}`);

      let imageBuffer = imageResult.status === 'fulfilled' ? imageResult.value : null;
      if (syncAnimatedImageToWordAudio && generateImage) {
        try {
          const noteInfoResult = (await client.notesInfo([noteId])) as StatsServerNoteInfo[];
          const noteInfo = noteInfoResult[0] ?? null;
          const animatedLeadInSeconds = noteInfo
            ? await resolveAnimatedImageLeadInSeconds({
                config: ankiConfig,
                noteInfo,
                resolveConfiguredFieldName: (candidateNoteInfo, ...preferredNames) =>
                  resolveStatsNoteFieldName(candidateNoteInfo, ...preferredNames),
                retrieveMediaFileBase64: (filename) => client.retrieveMediaFile(filename),
              })
            : 0;
          imageBuffer = await createImagePromise(animatedLeadInSeconds);
        } catch (err) {
          errors.push(`image: ${(err as Error).message}`);
        }
      }

      const mediaFields: Record<string, string> = {};
      const timestamp = Date.now();
      const sentenceFieldName = ankiConfig.fields?.sentence ?? 'Sentence';
      const audioFieldName = ankiConfig.fields?.audio ?? 'ExpressionAudio';
      const imageFieldName = ankiConfig.fields?.image ?? 'Picture';

      mediaFields[sentenceFieldName] = highlightedSentence;
      if (secondaryText) {
        mediaFields[ankiConfig.fields?.translation ?? 'SelectionText'] = secondaryText;
      }

      if (audioBuffer) {
        const audioFilename = `subminer_audio_${timestamp}.mp3`;
        try {
          await client.storeMediaFile(audioFilename, audioBuffer);
          mediaFields[audioFieldName] = `[sound:${audioFilename}]`;
        } catch (err) {
          errors.push(`audio upload: ${(err as Error).message}`);
        }
      }

      if (imageBuffer) {
        const imageExt = imageType === 'avif' ? 'avif' : (ankiConfig.media?.imageFormat ?? 'jpg');
        const imageFilename = `subminer_image_${timestamp}.${imageExt}`;
        try {
          await client.storeMediaFile(imageFilename, imageBuffer);
          mediaFields[imageFieldName] = `<img src="${imageFilename}">`;
        } catch (err) {
          errors.push(`image upload: ${(err as Error).message}`);
        }
      }

      const miscInfoFieldName = ankiConfig.fields?.miscInfo ?? '';
      if (miscInfoFieldName) {
        const pattern = ankiConfig.metadata?.pattern ?? '[SubMiner] %f (%t)';
        const filenameWithExt = videoTitle || basename(sourcePath);
        const filenameWithoutExt = filenameWithExt.replace(/\.[^.]+$/, '');
        const totalMs = Math.floor(startMs);
        const totalSec2 = Math.floor(totalMs / 1000);
        const hours = String(Math.floor(totalSec2 / 3600)).padStart(2, '0');
        const minutes = String(Math.floor((totalSec2 % 3600) / 60)).padStart(2, '0');
        const secs = String(totalSec2 % 60).padStart(2, '0');
        const ms = String(totalMs % 1000).padStart(3, '0');
        mediaFields[miscInfoFieldName] = pattern
          .replace(/%f/g, filenameWithoutExt)
          .replace(/%F/g, filenameWithExt)
          .replace(/%t/g, `${hours}:${minutes}:${secs}`)
          .replace(/%T/g, `${hours}:${minutes}:${secs}:${ms}`)
          .replace(/<br>/g, '\n');
      }

      if (Object.keys(mediaFields).length > 0) {
        try {
          await client.updateNoteFields(noteId, mediaFields);
        } catch (err) {
          errors.push(`update fields: ${(err as Error).message}`);
        }
      }

      return c.json({ noteId, ...(errors.length > 0 ? { errors } : {}) });
    }

    const [audioResult, imageResult] = await Promise.allSettled([audioPromise, imagePromise]);

    const audioBuffer = audioResult.status === 'fulfilled' ? audioResult.value : null;
    const imageBuffer = imageResult.status === 'fulfilled' ? imageResult.value : null;
    if (audioResult.status === 'rejected')
      errors.push(`audio: ${(audioResult.reason as Error).message}`);
    if (imageResult.status === 'rejected')
      errors.push(`image: ${(imageResult.reason as Error).message}`);

    const wordFieldName = getConfiguredWordFieldName(ankiConfig);
    const sentenceFieldName = ankiConfig.fields?.sentence ?? 'Sentence';
    const translationFieldName = ankiConfig.fields?.translation ?? 'SelectionText';
    const audioFieldName = ankiConfig.fields?.audio ?? 'ExpressionAudio';
    const imageFieldName = ankiConfig.fields?.image ?? 'Picture';
    const miscInfoFieldName = ankiConfig.fields?.miscInfo ?? '';

    const fields: Record<string, string> = {
      [sentenceFieldName]: highlightedSentence,
    };

    if (secondaryText) {
      fields[translationFieldName] = secondaryText;
    }

    if (ankiConfig.isLapis?.enabled || ankiConfig.isKiku?.enabled) {
      if (word) {
        fields[wordFieldName] = word;
      }
      if (mode === 'sentence') {
        fields['IsSentenceCard'] = 'x';
      } else if (mode === 'audio') {
        fields['IsAudioCard'] = 'x';
      }
    }

    const model = ankiConfig.isLapis?.sentenceCardModel || 'Basic';
    const deck = ankiConfig.deck ?? 'Default';
    const tags = ankiConfig.tags ?? ['SubMiner'];

    try {
      noteId = await client.addNote(deck, model, fields, tags);
    } catch (err) {
      return c.json({ error: `Failed to add note: ${(err as Error).message}` }, 502);
    }

    const mediaFields: Record<string, string> = {};
    const timestamp = Date.now();

    if (audioBuffer) {
      const audioFilename = `subminer_audio_${timestamp}.mp3`;
      try {
        await client.storeMediaFile(audioFilename, audioBuffer);
        mediaFields[audioFieldName] = `[sound:${audioFilename}]`;
      } catch (err) {
        errors.push(`audio upload: ${(err as Error).message}`);
      }
    }

    if (imageBuffer) {
      const imageExt = imageType === 'avif' ? 'avif' : (ankiConfig.media?.imageFormat ?? 'jpg');
      const imageFilename = `subminer_image_${timestamp}.${imageExt}`;
      try {
        await client.storeMediaFile(imageFilename, imageBuffer);
        mediaFields[imageFieldName] = `<img src="${imageFilename}">`;
      } catch (err) {
        errors.push(`image upload: ${(err as Error).message}`);
      }
    }

    if (miscInfoFieldName) {
      const pattern = ankiConfig.metadata?.pattern ?? '[SubMiner] %f (%t)';
      const filenameWithExt = videoTitle || basename(sourcePath);
      const filenameWithoutExt = filenameWithExt.replace(/\.[^.]+$/, '');
      const totalMs = Math.floor(startMs);
      const totalSec = Math.floor(totalMs / 1000);
      const hours = String(Math.floor(totalSec / 3600)).padStart(2, '0');
      const minutes = String(Math.floor((totalSec % 3600) / 60)).padStart(2, '0');
      const secs = String(totalSec % 60).padStart(2, '0');
      const ms = String(totalMs % 1000).padStart(3, '0');
      const miscInfo = pattern
        .replace(/%f/g, filenameWithoutExt)
        .replace(/%F/g, filenameWithExt)
        .replace(/%t/g, `${hours}:${minutes}:${secs}`)
        .replace(/%T/g, `${hours}:${minutes}:${secs}:${ms}`)
        .replace(/<br>/g, '\n');
      mediaFields[miscInfoFieldName] = miscInfo;
    }

    if (Object.keys(mediaFields).length > 0) {
      try {
        await client.updateNoteFields(noteId, mediaFields);
      } catch (err) {
        errors.push(`update fields: ${(err as Error).message}`);
      }
    }

    return c.json({ noteId, ...(errors.length > 0 ? { errors } : {}) });
  });

  if (options?.staticDir) {
    app.get('/assets/*', (c) => {
      const response = createStatsStaticResponse(options.staticDir!, c.req.path);
      if (!response) return c.text('Not found', 404);
      return response;
    });

    app.get('/index.html', (c) => {
      const response = createStatsStaticResponse(options.staticDir!, '/index.html');
      if (!response) return c.text('Stats UI not built', 404);
      return response;
    });

    app.get('*', (c) => {
      const staticResponse = createStatsStaticResponse(options.staticDir!, c.req.path);
      if (staticResponse) return staticResponse;
      const fallback = createStatsStaticResponse(options.staticDir!, '/index.html');
      if (!fallback) return c.text('Stats UI not built', 404);
      return fallback;
    });
  }

  return app;
}

export function startStatsServer(config: StatsServerConfig): { close: () => void } {
  const app = createStatsApp(config.tracker, {
    staticDir: config.staticDir,
    knownWordCachePath: config.knownWordCachePath,
    mpvSocketPath: config.mpvSocketPath,
    ankiConnectConfig: config.ankiConnectConfig,
    addYomitanNote: config.addYomitanNote,
    resolveAnkiNoteId: config.resolveAnkiNoteId,
  });

  const bunServe = (
    globalThis as typeof globalThis & {
      Bun: {
        serve: (options: {
          fetch: (typeof app)['fetch'];
          port: number;
          hostname: string;
        }) => { stop: () => void };
      };
    }
  ).Bun.serve;

  const server = bunServe({
    fetch: app.fetch,
    port: config.port,
    hostname: '127.0.0.1',
  });

  return {
    close: () => {
      server.stop();
    },
  };
}
