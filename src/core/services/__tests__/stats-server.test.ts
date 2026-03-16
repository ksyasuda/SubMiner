import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createStatsApp } from '../stats-server.js';
import type { ImmersionTrackerService } from '../immersion-tracker-service.js';

const SESSION_SUMMARIES = [
  {
    sessionId: 1,
    canonicalTitle: 'Test',
    videoId: 1,
    animeId: null,
    animeTitle: null,
    startedAtMs: Date.now(),
    endedAtMs: null,
    totalWatchedMs: 60_000,
    activeWatchedMs: 50_000,
    linesSeen: 10,
    wordsSeen: 100,
    tokensSeen: 80,
    cardsMined: 2,
    lookupCount: 5,
    lookupHits: 4,
  },
];

const DAILY_ROLLUPS = [
  {
    rollupDayOrMonth: Math.floor(Date.now() / 86_400_000),
    videoId: 1,
    totalSessions: 1,
    totalActiveMin: 10,
    totalLinesSeen: 10,
    totalWordsSeen: 100,
    totalTokensSeen: 80,
    totalCards: 2,
    cardsPerHour: 12,
    wordsPerMin: 10,
    lookupHitRate: 0.8,
  },
];

const VOCABULARY_STATS = [
  {
    wordId: 1,
    headword: 'する',
    word: 'する',
    reading: 'する',
    partOfSpeech: 'verb',
    pos1: '動詞',
    pos2: '自立',
    pos3: null,
    frequency: 100,
    frequencyRank: 42,
    animeCount: 2,
    firstSeen: Date.now(),
    lastSeen: Date.now(),
  },
];

const KANJI_STATS = [
  {
    kanjiId: 1,
    kanji: '日',
    frequency: 50,
    firstSeen: Date.now(),
    lastSeen: Date.now(),
  },
];

const OCCURRENCES = [
  {
    animeId: 1,
    animeTitle: 'Little Witch Academia',
    videoId: 2,
    videoTitle: 'Episode 4',
    sourcePath: '/media/anime/lwa/ep04.mkv',
    secondaryText: null,
    sessionId: 3,
    lineIndex: 7,
    segmentStartMs: 12_000,
    segmentEndMs: 14_500,
    text: '猫 猫 日 日 は 知っている',
    occurrenceCount: 2,
  },
];

const ANIME_LIBRARY = [
  {
    animeId: 1,
    canonicalTitle: 'Little Witch Academia',
    anilistId: 21858,
    totalSessions: 3,
    totalActiveMs: 180_000,
    totalCards: 5,
    totalWordsSeen: 300,
    episodeCount: 2,
    episodesTotal: 25,
    lastWatchedMs: Date.now(),
  },
];

const ANIME_DETAIL = {
  animeId: 1,
  canonicalTitle: 'Little Witch Academia',
  anilistId: 21858,
  titleRomaji: 'Little Witch Academia',
  titleEnglish: 'Little Witch Academia',
  titleNative: 'リトルウィッチアカデミア',
  totalSessions: 3,
  totalActiveMs: 180_000,
  totalCards: 5,
  totalWordsSeen: 300,
  totalLinesSeen: 50,
  totalLookupCount: 20,
  totalLookupHits: 15,
  episodeCount: 2,
  lastWatchedMs: Date.now(),
};

const ANIME_WORDS = [
  {
    wordId: 1,
    headword: '魔法',
    word: '魔法',
    reading: 'まほう',
    partOfSpeech: 'noun',
    frequency: 42,
  },
];

const EPISODES_PER_DAY = [
  { epochDay: Math.floor(Date.now() / 86_400_000) - 1, episodeCount: 3 },
  { epochDay: Math.floor(Date.now() / 86_400_000), episodeCount: 1 },
];

const NEW_ANIME_PER_DAY = [{ epochDay: Math.floor(Date.now() / 86_400_000) - 2, newAnimeCount: 2 }];

const WATCH_TIME_PER_ANIME = [
  {
    epochDay: Math.floor(Date.now() / 86_400_000) - 1,
    animeId: 1,
    animeTitle: 'Little Witch Academia',
    totalActiveMin: 25,
  },
];

const ANIME_EPISODES = [
  {
    animeId: 1,
    videoId: 1,
    canonicalTitle: 'Episode 1',
    parsedTitle: 'Little Witch Academia',
    season: 1,
    episode: 1,
    totalSessions: 1,
    totalActiveMs: 90_000,
    totalCards: 3,
    totalWordsSeen: 150,
    lastWatchedMs: Date.now(),
  },
];

const WORD_DETAIL = {
  wordId: 1,
  headword: '猫',
  word: '猫',
  reading: 'ねこ',
  partOfSpeech: 'noun',
  pos1: '名詞',
  pos2: '一般',
  pos3: null,
  frequency: 42,
  firstSeen: Date.now() - 100_000,
  lastSeen: Date.now(),
};

const WORD_ANIME_APPEARANCES = [
  { animeId: 1, animeTitle: 'Little Witch Academia', occurrenceCount: 12 },
];

const SIMILAR_WORDS = [
  { wordId: 2, headword: '猫耳', word: '猫耳', reading: 'ねこみみ', frequency: 5 },
];

const KANJI_DETAIL = {
  kanjiId: 1,
  kanji: '日',
  frequency: 50,
  firstSeen: Date.now() - 100_000,
  lastSeen: Date.now(),
};

const KANJI_ANIME_APPEARANCES = [
  { animeId: 1, animeTitle: 'Little Witch Academia', occurrenceCount: 30 },
];

const KANJI_WORDS = [
  { wordId: 3, headword: '日本', word: '日本', reading: 'にほん', frequency: 20 },
];

const EPISODE_CARD_EVENTS = [
  { eventId: 1, sessionId: 1, tsMs: Date.now(), cardsDelta: 1, noteIds: [12345] },
];

function createMockTracker(
  overrides: Partial<ImmersionTrackerService> = {},
): ImmersionTrackerService {
  return {
    getSessionSummaries: async () => SESSION_SUMMARIES,
    getDailyRollups: async () => DAILY_ROLLUPS,
    getMonthlyRollups: async () => [],
    getQueryHints: async () => ({
      totalSessions: 5,
      activeSessions: 1,
      episodesToday: 2,
      activeAnimeCount: 3,
    }),
    getSessionTimeline: async () => [],
    getSessionEvents: async () => [],
    getVocabularyStats: async () => VOCABULARY_STATS,
    getKanjiStats: async () => KANJI_STATS,
    getWordOccurrences: async () => OCCURRENCES,
    getKanjiOccurrences: async () => OCCURRENCES,
    getAnimeLibrary: async () => ANIME_LIBRARY,
    getAnimeDetail: async (animeId: number) => (animeId === 1 ? ANIME_DETAIL : null),
    getAnimeEpisodes: async () => ANIME_EPISODES,
    getAnimeAnilistEntries: async () => [],
    getAnimeWords: async () => ANIME_WORDS,
    getAnimeDailyRollups: async () => DAILY_ROLLUPS,
    getEpisodesPerDay: async () => EPISODES_PER_DAY,
    getNewAnimePerDay: async () => NEW_ANIME_PER_DAY,
    getWatchTimePerAnime: async () => WATCH_TIME_PER_ANIME,
    getStreakCalendar: async () => [
      { epochDay: Math.floor(Date.now() / 86_400_000) - 1, totalActiveMin: 30 },
      { epochDay: Math.floor(Date.now() / 86_400_000), totalActiveMin: 45 },
    ],
    getAnimeCoverArt: async (animeId: number) =>
      animeId === 1
        ? {
            videoId: 1,
            anilistId: 21858,
            coverUrl: 'https://example.com/cover.jpg',
            coverBlob: Buffer.from([0xff, 0xd8, 0xff, 0xd9]),
            titleRomaji: 'Little Witch Academia',
            titleEnglish: 'Little Witch Academia',
            episodesTotal: 25,
            fetchedAtMs: Date.now(),
          }
        : null,
    getWordDetail: async (wordId: number) => (wordId === 1 ? WORD_DETAIL : null),
    getWordAnimeAppearances: async () => WORD_ANIME_APPEARANCES,
    getSimilarWords: async () => SIMILAR_WORDS,
    getKanjiDetail: async (kanjiId: number) => (kanjiId === 1 ? KANJI_DETAIL : null),
    getKanjiAnimeAppearances: async () => KANJI_ANIME_APPEARANCES,
    getKanjiWords: async () => KANJI_WORDS,
    getEpisodeWords: async () => ANIME_WORDS,
    getEpisodeSessions: async () => SESSION_SUMMARIES,
    getEpisodeCardEvents: async () => EPISODE_CARD_EVENTS,
    ...overrides,
  } as unknown as ImmersionTrackerService;
}

function withTempDir<T>(fn: (dir: string) => Promise<T> | T): Promise<T> | T {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-stats-server-test-'));
  const result = fn(dir);
  if (result instanceof Promise) {
    return result.finally(() => {
      fs.rmSync(dir, { recursive: true, force: true });
    });
  }
  fs.rmSync(dir, { recursive: true, force: true });
  return result;
}

describe('stats server API routes', () => {
  it('GET /api/stats/overview returns overview data', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/overview');
    assert.equal(res.status, 200);
    assert.equal(res.headers.get('access-control-allow-origin'), null);
    const body = await res.json();
    assert.ok(body.sessions);
    assert.ok(body.rollups);
    assert.ok(body.hints);
    assert.equal(body.hints.totalSessions, 5);
    assert.equal(body.hints.activeSessions, 1);
    assert.equal(body.hints.episodesToday, 2);
    assert.equal(body.hints.activeAnimeCount, 3);
  });

  it('GET /api/stats/sessions returns session list', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/sessions?limit=5');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
  });

  it('GET /api/stats/vocabulary returns word frequency data', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/vocabulary');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body[0].headword, 'する');
  });

  it('GET /api/stats/kanji returns kanji frequency data', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/kanji');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body[0].kanji, '日');
  });

  it('GET /api/stats/streak-calendar returns streak calendar rows', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/streak-calendar');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body.length, 2);
    assert.equal(body[0].totalActiveMin, 30);
    assert.equal(body[1].totalActiveMin, 45);
  });

  it('GET /api/stats/streak-calendar clamps oversized days', async () => {
    let seenDays = 0;
    const app = createStatsApp(
      createMockTracker({
        getStreakCalendar: async (days?: number) => {
          seenDays = days ?? 0;
          return [];
        },
      }),
    );

    const res = await app.request('/api/stats/streak-calendar?days=999999');
    assert.equal(res.status, 200);
    assert.equal(seenDays, 365);
  });

  it('GET /api/stats/trends/episodes-per-day returns episode count rows', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/trends/episodes-per-day');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body.length, 2);
    assert.equal(body[0].episodeCount, 3);
  });

  it('GET /api/stats/trends/episodes-per-day clamps oversized limits', async () => {
    let seenLimit = 0;
    const app = createStatsApp(
      createMockTracker({
        getEpisodesPerDay: async (limit?: number) => {
          seenLimit = limit ?? 0;
          return EPISODES_PER_DAY;
        },
      }),
    );
    const res = await app.request('/api/stats/trends/episodes-per-day?limit=999999');
    assert.equal(res.status, 200);
    assert.equal(seenLimit, 365);
  });

  it('GET /api/stats/trends/new-anime-per-day returns new anime rows', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/trends/new-anime-per-day');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body.length, 1);
    assert.equal(body[0].newAnimeCount, 2);
  });

  it('GET /api/stats/trends/new-anime-per-day clamps oversized limits', async () => {
    let seenLimit = 0;
    const app = createStatsApp(
      createMockTracker({
        getNewAnimePerDay: async (limit?: number) => {
          seenLimit = limit ?? 0;
          return NEW_ANIME_PER_DAY;
        },
      }),
    );
    const res = await app.request('/api/stats/trends/new-anime-per-day?limit=999999');
    assert.equal(res.status, 200);
    assert.equal(seenLimit, 365);
  });

  it('GET /api/stats/trends/watch-time-per-anime returns watch time rows', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/trends/watch-time-per-anime');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body.length, 1);
    assert.equal(body[0].animeTitle, 'Little Witch Academia');
    assert.equal(body[0].totalActiveMin, 25);
  });

  it('GET /api/stats/trends/watch-time-per-anime clamps oversized limits', async () => {
    let seenLimit = 0;
    const app = createStatsApp(
      createMockTracker({
        getWatchTimePerAnime: async (limit?: number) => {
          seenLimit = limit ?? 0;
          return WATCH_TIME_PER_ANIME;
        },
      }),
    );
    const res = await app.request('/api/stats/trends/watch-time-per-anime?limit=999999');
    assert.equal(res.status, 200);
    assert.equal(seenLimit, 365);
  });

  it('GET /api/stats/vocabulary/occurrences returns recent occurrence rows for a word', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getWordOccurrences: async (...args: unknown[]) => {
          seenArgs = args;
          return OCCURRENCES;
        },
      }),
    );

    const res = await app.request(
      '/api/stats/vocabulary/occurrences?headword=%E7%8C%AB&word=%E7%8C%AB&reading=%E3%81%AD%E3%81%93&limit=999999&offset=25',
    );
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body[0].animeTitle, 'Little Witch Academia');
    assert.deepEqual(seenArgs, ['猫', '猫', 'ねこ', 500, 25]);
  });

  it('GET /api/stats/kanji/occurrences returns recent occurrence rows for a kanji', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getKanjiOccurrences: async (...args: unknown[]) => {
          seenArgs = args;
          return OCCURRENCES;
        },
      }),
    );

    const res = await app.request(
      '/api/stats/kanji/occurrences?kanji=%E6%97%A5&limit=999999&offset=10',
    );
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body[0].occurrenceCount, 2);
    assert.deepEqual(seenArgs, ['日', 500, 10]);
  });

  it('GET /api/stats/vocabulary/occurrences rejects missing required params', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/vocabulary/occurrences?headword=%E7%8C%AB');
    assert.equal(res.status, 400);
  });

  it('GET /api/stats/vocabulary clamps oversized limits', async () => {
    let seenLimit = 0;
    const app = createStatsApp(
      createMockTracker({
        getVocabularyStats: async (limit?: number, _excludePos?: string[]) => {
          seenLimit = limit ?? 0;
          return VOCABULARY_STATS;
        },
      }),
    );

    const res = await app.request('/api/stats/vocabulary?limit=999999');
    assert.equal(res.status, 200);
    assert.equal(seenLimit, 500);
  });

  it('GET /api/stats/vocabulary passes excludePos to tracker', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getVocabularyStats: async (...args: unknown[]) => {
          seenArgs = args;
          return VOCABULARY_STATS;
        },
      }),
    );

    const res = await app.request('/api/stats/vocabulary?excludePos=particle,auxiliary');
    assert.equal(res.status, 200);
    assert.deepEqual(seenArgs, [100, ['particle', 'auxiliary']]);
  });

  it('GET /api/stats/vocabulary returns POS fields', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/vocabulary');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.equal(body[0].partOfSpeech, 'verb');
    assert.equal(body[0].pos1, '動詞');
    assert.equal(body[0].pos2, '自立');
    assert.equal(body[0].pos3, null);
  });

  it('GET /api/stats/anime returns anime library', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/anime');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body[0].canonicalTitle, 'Little Witch Academia');
  });

  it('GET /api/stats/anime/:animeId returns anime detail with episodes', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/anime/1');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(body.detail);
    assert.equal(body.detail.canonicalTitle, 'Little Witch Academia');
    assert.ok(Array.isArray(body.episodes));
    assert.equal(body.episodes[0].videoId, 1);
  });

  it('GET /api/stats/anime/:animeId returns 404 for missing anime', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/anime/99999');
    assert.equal(res.status, 404);
  });

  it('GET /api/stats/anime/:animeId/cover returns cover art', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/anime/1/cover');
    assert.equal(res.status, 200);
    assert.equal(res.headers.get('content-type'), 'image/jpeg');
    assert.equal(res.headers.get('cache-control'), 'public, max-age=86400');
  });

  it('GET /api/stats/anime/:animeId/cover returns 404 for missing anime', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/anime/99999/cover');
    assert.equal(res.status, 404);
  });

  it('GET /api/stats/anime/:animeId/words returns top words for an anime', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getAnimeWords: async (...args: unknown[]) => {
          seenArgs = args;
          return ANIME_WORDS;
        },
      }),
    );

    const res = await app.request('/api/stats/anime/1/words?limit=25');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body[0].headword, '魔法');
    assert.deepEqual(seenArgs, [1, 25]);
  });

  it('GET /api/stats/anime/:animeId/words rejects invalid animeId', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/anime/0/words');
    assert.equal(res.status, 400);
  });

  it('GET /api/stats/anime/:animeId/words clamps oversized limits', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getAnimeWords: async (...args: unknown[]) => {
          seenArgs = args;
          return ANIME_WORDS;
        },
      }),
    );

    const res = await app.request('/api/stats/anime/1/words?limit=999999');
    assert.equal(res.status, 200);
    assert.deepEqual(seenArgs, [1, 200]);
  });

  it('GET /api/stats/anime/:animeId/rollups returns daily rollups for an anime', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getAnimeDailyRollups: async (...args: unknown[]) => {
          seenArgs = args;
          return DAILY_ROLLUPS;
        },
      }),
    );

    const res = await app.request('/api/stats/anime/1/rollups?limit=30');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
    assert.equal(body[0].totalSessions, 1);
    assert.deepEqual(seenArgs, [1, 30]);
  });

  it('GET /api/stats/anime/:animeId/rollups rejects invalid animeId', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/anime/-1/rollups');
    assert.equal(res.status, 400);
  });

  it('GET /api/stats/anime/:animeId/rollups clamps oversized limits', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getAnimeDailyRollups: async (...args: unknown[]) => {
          seenArgs = args;
          return DAILY_ROLLUPS;
        },
      }),
    );

    const res = await app.request('/api/stats/anime/1/rollups?limit=999999');
    assert.equal(res.status, 200);
    assert.deepEqual(seenArgs, [1, 365]);
  });

  it('GET /api/stats/vocabulary/:wordId/detail returns word detail', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/vocabulary/1/detail');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(body.detail);
    assert.equal(body.detail.headword, '猫');
    assert.equal(body.detail.wordId, 1);
    assert.ok(Array.isArray(body.animeAppearances));
    assert.equal(body.animeAppearances[0].animeTitle, 'Little Witch Academia');
    assert.ok(Array.isArray(body.similarWords));
    assert.equal(body.similarWords[0].headword, '猫耳');
  });

  it('GET /api/stats/vocabulary/:wordId/detail returns 404 for missing word', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/vocabulary/99999/detail');
    assert.equal(res.status, 404);
  });

  it('GET /api/stats/vocabulary/:wordId/detail returns 400 for invalid wordId', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/vocabulary/0/detail');
    assert.equal(res.status, 400);
  });

  it('GET /api/stats/kanji/:kanjiId/detail returns kanji detail', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/kanji/1/detail');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(body.detail);
    assert.equal(body.detail.kanji, '日');
    assert.equal(body.detail.kanjiId, 1);
    assert.ok(Array.isArray(body.animeAppearances));
    assert.equal(body.animeAppearances[0].animeTitle, 'Little Witch Academia');
    assert.ok(Array.isArray(body.words));
    assert.equal(body.words[0].headword, '日本');
  });

  it('GET /api/stats/kanji/:kanjiId/detail returns 404 for missing kanji', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/kanji/99999/detail');
    assert.equal(res.status, 404);
  });

  it('GET /api/stats/kanji/:kanjiId/detail returns 400 for invalid kanjiId', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/kanji/0/detail');
    assert.equal(res.status, 400);
  });

  it('GET /api/stats/vocabulary/occurrences still works with detail routes present', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request(
      '/api/stats/vocabulary/occurrences?headword=%E7%8C%AB&word=%E7%8C%AB&reading=%E3%81%AD%E3%81%93',
    );
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
  });

  it('GET /api/stats/kanji/occurrences still works with detail routes present', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/kanji/occurrences?kanji=%E6%97%A5');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
  });

  it('GET /api/stats/episode/:videoId/detail returns episode detail', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/episode/1/detail');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body.sessions));
    assert.ok(Array.isArray(body.words));
    assert.ok(Array.isArray(body.cardEvents));
    assert.equal(body.cardEvents[0].noteIds[0], 12345);
  });

  it('GET /api/stats/episode/:videoId/detail returns 400 for invalid videoId', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/episode/0/detail');
    assert.equal(res.status, 400);
  });

  it('DELETE /api/stats/sessions/:sessionId deletes a session', async () => {
    let deletedSessionId = 0;
    const app = createStatsApp(
      createMockTracker({
        deleteSession: async (sessionId: number) => {
          deletedSessionId = sessionId;
        },
      }),
    );

    const res = await app.request('/api/stats/sessions/42', { method: 'DELETE' });

    assert.equal(res.status, 200);
    assert.equal(deletedSessionId, 42);
    assert.deepEqual(await res.json(), { ok: true });
  });

  it('POST /api/stats/anki/browse returns 400 for missing noteId', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/anki/browse', { method: 'POST' });
    assert.equal(res.status, 400);
  });

  it('serves stats index and asset files from absolute static dir paths', async () => {
    await withTempDir(async (dir) => {
      const assetDir = path.join(dir, 'assets');
      fs.mkdirSync(assetDir, { recursive: true });
      fs.writeFileSync(
        path.join(dir, 'index.html'),
        '<!doctype html><html><body><div id="root"></div><script src="./assets/app.js"></script></body></html>',
      );
      fs.writeFileSync(path.join(assetDir, 'app.js'), 'console.log("stats ok");');

      const app = createStatsApp(createMockTracker(), { staticDir: dir });
      const indexRes = await app.request('/');
      assert.equal(indexRes.status, 200);
      assert.match(await indexRes.text(), /assets\/app\.js/);

      const assetRes = await app.request('/assets/app.js');
      assert.equal(assetRes.status, 200);
      assert.equal(assetRes.headers.get('content-type'), 'text/javascript; charset=utf-8');
      assert.match(await assetRes.text(), /stats ok/);
    });
  });

  it('fetches and serves missing cover art on demand', async () => {
    let ensureCalls = 0;
    let hasCover = false;
    const app = createStatsApp(
      createMockTracker({
        getCoverArt: async () =>
          hasCover
            ? {
                videoId: 1,
                anilistId: 1,
                coverUrl: 'https://example.com/cover.jpg',
                coverBlob: Buffer.from([0xff, 0xd8, 0xff, 0xd9]),
                titleRomaji: 'Test',
                titleEnglish: 'Test',
                episodesTotal: 12,
                fetchedAtMs: Date.now(),
              }
            : null,
        ensureCoverArt: async () => {
          ensureCalls += 1;
          hasCover = true;
          return true;
        },
      }),
    );

    const res = await app.request('/api/stats/media/1/cover');
    assert.equal(res.status, 200);
    assert.equal(res.headers.get('content-type'), 'image/jpeg');
    assert.equal(ensureCalls, 1);
  });
});
