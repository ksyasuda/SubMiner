import type { Hono } from 'hono';
import { statsJson } from '../../../types/stats-http-contract.js';
import type { ImmersionTrackerService } from '../immersion-tracker-service.js';
import {
  buildSentenceSearchOptions,
  enrichSessionsWithKnownWordMetrics,
  parseBooleanQuery,
  parseExcludedWordsBody,
  parseIntQuery,
} from './route-support.js';

export function registerStatsLibraryRoutes(
  app: Hono,
  tracker: ImmersionTrackerService,
  options?: {
    knownWordCachePath?: string;
    resolveSentenceSearchHeadwords?: (term: string) => Promise<string[]> | string[];
  },
): void {
  app.get('/api/stats/vocabulary', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 100, 500);
    const excludePos = c.req.query('excludePos')?.split(',').filter(Boolean);
    const vocab = await tracker.getVocabularyStats(limit, excludePos);
    return c.json(statsJson('vocabulary', vocab));
  });

  app.get('/api/stats/excluded-words', async (c) => {
    return c.json(statsJson('excludedWords', await tracker.getStatsExcludedWords()));
  });

  app.put('/api/stats/excluded-words', async (c) => {
    const body = await c.req.json().catch(() => null);
    const words = parseExcludedWordsBody(body);
    if (!words) return c.body(null, 400);
    await tracker.replaceStatsExcludedWords(words);
    return c.json(statsJson('setExcludedWords', { ok: true }));
  });

  app.get('/api/stats/vocabulary/occurrences', async (c) => {
    const headword = (c.req.query('headword') ?? '').trim();
    const word = (c.req.query('word') ?? '').trim();
    const reading = (c.req.query('reading') ?? '').trim();
    if (!headword || !word) {
      return c.json(statsJson('wordOccurrences', []), 400);
    }
    const limit = parseIntQuery(c.req.query('limit'), 50, 500);
    const offset = parseIntQuery(c.req.query('offset'), 0, 10_000);
    const occurrences = await tracker.getWordOccurrences(headword, word, reading, limit, offset);
    return c.json(statsJson('wordOccurrences', occurrences));
  });

  app.get('/api/stats/sentences/search', async (c) => {
    const query = (c.req.query('q') ?? '').trim();
    if (!query) return c.json(statsJson('sentenceSearch', []));
    const limit = parseIntQuery(c.req.query('limit'), 50, 100);
    const searchByHeadword = parseBooleanQuery(c.req.query('headword'), true);
    const searchOptions = await buildSentenceSearchOptions(
      query,
      searchByHeadword,
      options?.resolveSentenceSearchHeadwords,
    );
    const rows = await tracker.searchSubtitleSentences(query, limit, searchOptions);
    return c.json(statsJson('sentenceSearch', rows));
  });

  app.get('/api/stats/kanji', async (c) => {
    const limit = parseIntQuery(c.req.query('limit'), 100, 500);
    const kanji = await tracker.getKanjiStats(limit);
    return c.json(statsJson('kanji', kanji));
  });

  app.get('/api/stats/kanji/occurrences', async (c) => {
    const kanji = (c.req.query('kanji') ?? '').trim();
    if (!kanji) {
      return c.json(statsJson('kanjiOccurrences', []), 400);
    }
    const limit = parseIntQuery(c.req.query('limit'), 50, 500);
    const offset = parseIntQuery(c.req.query('offset'), 0, 10_000);
    const occurrences = await tracker.getKanjiOccurrences(kanji, limit, offset);
    return c.json(statsJson('kanjiOccurrences', occurrences));
  });

  app.get('/api/stats/vocabulary/:wordId/detail', async (c) => {
    const wordId = parseIntQuery(c.req.param('wordId'), 0);
    if (wordId <= 0) return c.body(null, 400);
    const detail = await tracker.getWordDetail(wordId);
    if (!detail) return c.body(null, 404);
    const animeAppearances = await tracker.getWordAnimeAppearances(wordId);
    const similarWords = await tracker.getSimilarWords(wordId);
    return c.json(statsJson('wordDetail', { detail, animeAppearances, similarWords }));
  });

  app.get('/api/stats/kanji/:kanjiId/detail', async (c) => {
    const kanjiId = parseIntQuery(c.req.param('kanjiId'), 0);
    if (kanjiId <= 0) return c.body(null, 400);
    const detail = await tracker.getKanjiDetail(kanjiId);
    if (!detail) return c.body(null, 404);
    const animeAppearances = await tracker.getKanjiAnimeAppearances(kanjiId);
    const words = await tracker.getKanjiWords(kanjiId);
    return c.json(statsJson('kanjiDetail', { detail, animeAppearances, words }));
  });

  app.get('/api/stats/media', async (c) => {
    const library = await tracker.getMediaLibrary();
    return c.json(statsJson('mediaLibrary', library));
  });

  app.get('/api/stats/media/:videoId', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.json(statsJson('error', null), 400);
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
    return c.json(statsJson('mediaDetail', { detail, sessions, rollups }));
  });

  app.get('/api/stats/anime', async (c) => {
    const rows = await tracker.getAnimeLibrary();
    return c.json(statsJson('animeLibrary', rows));
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
    return c.json(statsJson('animeDetail', { detail, episodes, anilistEntries }));
  });

  app.get('/api/stats/anime/:animeId/words', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    const limit = parseIntQuery(c.req.query('limit'), 50, 200);
    if (animeId <= 0) return c.body(null, 400);
    return c.json(statsJson('animeWords', await tracker.getAnimeWords(animeId, limit)));
  });

  app.get('/api/stats/anime/:animeId/rollups', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    const limit = parseIntQuery(c.req.query('limit'), 90, 365);
    if (animeId <= 0) return c.body(null, 400);
    return c.json(statsJson('animeRollups', await tracker.getAnimeDailyRollups(animeId, limit)));
  });

  app.patch('/api/stats/media/:videoId/watched', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.body(null, 400);
    const body = await c.req.json().catch(() => null);
    const watched = typeof body?.watched === 'boolean' ? body.watched : true;
    await tracker.setVideoWatched(videoId, watched);
    return c.json(statsJson('setVideoWatched', { ok: true }));
  });

  app.delete('/api/stats/sessions', async (c) => {
    const body = await c.req.json().catch(() => null);
    const ids = Array.isArray(body?.sessionIds)
      ? body.sessionIds.filter((id: unknown) => typeof id === 'number' && id > 0)
      : [];
    if (ids.length === 0) return c.body(null, 400);
    await tracker.deleteSessions(ids);
    return c.json(statsJson('deleteSessions', { ok: true }));
  });

  app.delete('/api/stats/sessions/:sessionId', async (c) => {
    const sessionId = parseIntQuery(c.req.param('sessionId'), 0);
    if (sessionId <= 0) return c.body(null, 400);
    await tracker.deleteSession(sessionId);
    return c.json(statsJson('deleteSession', { ok: true }));
  });

  app.delete('/api/stats/media/:videoId', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.body(null, 400);
    await tracker.deleteVideo(videoId);
    return c.json(statsJson('deleteVideo', { ok: true }));
  });
}
