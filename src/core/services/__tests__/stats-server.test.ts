import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import type { AddressInfo } from 'node:net';
import { createStatsApp, startStatsServer } from '../stats-server.js';
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
    tokensSeen: 80,
    cardsMined: 2,
    lookupCount: 5,
    lookupHits: 4,
    yomitanLookupCount: 5,
  },
];

const DAILY_ROLLUPS = [
  {
    rollupDayOrMonth: Math.floor(Date.now() / 86_400_000),
    videoId: 1,
    totalSessions: 1,
    totalActiveMin: 10,
    totalLinesSeen: 10,
    totalTokensSeen: 80,
    totalCards: 2,
    cardsPerHour: 12,
    tokensPerMin: 10,
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
    totalTokensSeen: 300,
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
  totalTokensSeen: 300,
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

const TRENDS_DASHBOARD = {
  activity: {
    watchTime: [{ label: 'Mar 1', value: 25 }],
    cards: [{ label: 'Mar 1', value: 5 }],
    words: [{ label: 'Mar 1', value: 300 }],
    sessions: [{ label: 'Mar 1', value: 3 }],
  },
  progress: {
    watchTime: [{ label: 'Mar 1', value: 25 }],
    sessions: [{ label: 'Mar 1', value: 3 }],
    words: [{ label: 'Mar 1', value: 300 }],
    newWords: [{ label: 'Mar 1', value: 12 }],
    cards: [{ label: 'Mar 1', value: 5 }],
    episodes: [{ label: 'Mar 1', value: 2 }],
    lookups: [{ label: 'Mar 1', value: 15 }],
  },
  ratios: {
    lookupsPerHundred: [{ label: 'Mar 1', value: 5 }],
  },
  librarySummary: [
    {
      title: 'Little Witch Academia',
      watchTimeMin: 25,
      videos: 1,
      sessions: 1,
      cards: 5,
      words: 300,
      lookups: 15,
      lookupsPerHundred: 5,
      firstWatched: 20_000,
      lastWatched: 20_000,
    },
  ],
  animeCumulative: {
    watchTime: [{ epochDay: 20_000, animeTitle: 'Little Witch Academia', value: 25 }],
    episodes: [{ epochDay: 20_000, animeTitle: 'Little Witch Academia', value: 1 }],
    cards: [{ epochDay: 20_000, animeTitle: 'Little Witch Academia', value: 5 }],
    words: [{ epochDay: 20_000, animeTitle: 'Little Witch Academia', value: 300 }],
  },
  patterns: {
    watchTimeByDayOfWeek: [{ label: 'Sun', value: 25 }],
    watchTimeByHour: [{ label: '12:00', value: 25 }],
  },
};

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
    totalTokensSeen: 150,
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
      totalEpisodesWatched: 0,
      totalAnimeCompleted: 0,
      totalActiveMin: 120,
      totalCards: 0,
      activeDays: 7,
      totalTokensSeen: 80,
      totalLookupCount: 5,
      totalLookupHits: 4,
      totalYomitanLookupCount: 5,
      newWordsToday: 0,
      newWordsThisWeek: 0,
    }),
    getSessionTimeline: async () => [],
    getSessionEvents: async () => [],
    getVocabularyStats: async () => VOCABULARY_STATS,
    getStatsExcludedWords: async () => [],
    replaceStatsExcludedWords: async () => {},
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
    getTrendsDashboard: async () => TRENDS_DASHBOARD,
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

type CapturedAnkiRequest = {
  action?: string;
  params?: {
    cards?: number[];
    deck?: string;
    notes?: number[];
    query?: string;
    note?: {
      id?: number;
      deckName?: string;
      modelName?: string;
      fields?: Record<string, string>;
      tags?: string[];
    };
  };
};

async function withFakeAnkiConnect<T>(
  fn: (requests: CapturedAnkiRequest[], url: string) => Promise<T>,
): Promise<T> {
  const requests: CapturedAnkiRequest[] = [];
  const server = http.createServer((req, res) => {
    const chunks: Buffer[] = [];
    req.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
    req.on('end', () => {
      const payload = JSON.parse(Buffer.concat(chunks).toString('utf8')) as CapturedAnkiRequest;
      requests.push(payload);

      let body: unknown = { result: null, error: null };
      if (payload.action === 'addNote') {
        if (!payload.params?.note?.deckName) {
          body = { result: null, error: 'deck was not found: ' };
        } else {
          body = { result: 12345, error: null };
        }
      } else if (payload.action === 'notesInfo') {
        const noteIds = payload.params?.notes ?? [];
        body = {
          result: noteIds.map((noteId) => ({
            noteId,
            fields: {
              Expression: { value: '猫' },
              ExpressionAudio: { value: '[sound:word.mp3]' },
              Sentence: { value: '' },
              SentenceAudio: { value: '' },
              Picture: { value: '' },
              MiscInfo: { value: '' },
              SelectionText: { value: '' },
            },
          })),
          error: null,
        };
      } else if (payload.action === 'findCards') {
        body = { result: [9001], error: null };
      } else if (payload.action === 'changeDeck') {
        body = { result: null, error: null };
      }

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(body));
    });
  });

  await new Promise<void>((resolve) => {
    server.listen(0, '127.0.0.1', resolve);
  });

  try {
    const address = server.address() as AddressInfo;
    return await fn(requests, `http://127.0.0.1:${address.port}`);
  } finally {
    await new Promise<void>((resolve, reject) => {
      server.close((err) => (err ? reject(err) : resolve()));
    });
  }
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
    assert.equal(body.hints.totalEpisodesWatched, 0);
    assert.equal(body.hints.totalAnimeCompleted, 0);
    assert.equal(body.hints.totalActiveMin, 120);
    assert.equal(body.hints.activeDays, 7);
    assert.equal(body.hints.totalTokensSeen, 80);
    assert.equal(body.hints.totalYomitanLookupCount, 5);
  });

  it('GET /api/stats/sessions returns session list', async () => {
    const app = createStatsApp(createMockTracker());
    const res = await app.request('/api/stats/sessions?limit=5');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.ok(Array.isArray(body));
  });

  it('GET /api/stats/sentences/search resolves headword candidates by default', async () => {
    const seen: Array<{
      query: string;
      limit: number;
      options: unknown;
    }> = [];
    const resolvedTerms: string[] = [];
    const app = createStatsApp(
      createMockTracker({
        searchSubtitleSentences: async (query: string, limit: number, options: unknown) => {
          seen.push({ query, limit, options });
          return [];
        },
      }),
      {
        resolveSentenceSearchHeadwords: async (term: string) => {
          resolvedTerms.push(term);
          return term === '知らない' ? ['知る'] : [];
        },
      },
    );

    const res = await app.request(
      '/api/stats/sentences/search?q=%E7%9F%A5%E3%82%89%E3%81%AA%E3%81%84&limit=12',
    );

    assert.equal(res.status, 200);
    assert.deepEqual(resolvedTerms, ['知らない']);
    assert.deepEqual(seen, [
      {
        query: '知らない',
        limit: 12,
        options: { headwordTerms: [{ term: '知らない', headwords: ['知る'] }] },
      },
    ]);
  });

  it('GET /api/stats/sessions enriches known-word metrics using filtered persisted totals', async () => {
    await withTempDir(async (dir) => {
      const cachePath = path.join(dir, 'known-words.json');
      fs.writeFileSync(
        cachePath,
        JSON.stringify({
          version: 1,
          words: ['する'],
        }),
      );

      const app = createStatsApp(
        createMockTracker({
          getSessionWordsByLine: async (sessionId: number) =>
            sessionId === 1
              ? [
                  { lineIndex: 1, headword: 'する', occurrenceCount: 2 },
                  { lineIndex: 2, headword: '未知', occurrenceCount: 1 },
                ]
              : [],
        }),
        { knownWordCachePath: cachePath },
      );

      const res = await app.request('/api/stats/sessions?limit=5');
      assert.equal(res.status, 200);
      const body = await res.json();
      const first = body[0];
      assert.equal(first.knownWordsSeen, 2);
      assert.equal(first.knownWordRate, 66.7);
    });
  });

  it('GET /api/stats/sessions/:id/events forwards event type filters to the tracker', async () => {
    let seenSessionId = 0;
    let seenLimit = 0;
    let seenTypes: number[] | undefined;
    const app = createStatsApp(
      createMockTracker({
        getSessionEvents: async (sessionId: number, limit?: number, eventTypes?: number[]) => {
          seenSessionId = sessionId;
          seenLimit = limit ?? 0;
          seenTypes = eventTypes;
          return [];
        },
      }),
    );

    const res = await app.request('/api/stats/sessions/7/events?limit=12&types=4,5,9');
    assert.equal(res.status, 200);
    assert.equal(seenSessionId, 7);
    assert.equal(seenLimit, 12);
    assert.deepEqual(seenTypes, [4, 5, 9]);
  });

  it('GET /api/stats/sessions/:id/timeline requests the full session when no limit is provided', async () => {
    let seenSessionId = 0;
    let seenLimit: number | undefined;
    const app = createStatsApp(
      createMockTracker({
        getSessionTimeline: async (sessionId: number, limit?: number) => {
          seenSessionId = sessionId;
          seenLimit = limit;
          return [];
        },
      }),
    );

    const res = await app.request('/api/stats/sessions/7/timeline');
    assert.equal(res.status, 200);
    assert.equal(seenSessionId, 7);
    assert.equal(seenLimit, undefined);
  });

  it('GET /api/stats/sessions/:id/known-words-timeline preserves line positions and counts filtered totals', async () => {
    await withTempDir(async (dir) => {
      const cachePath = path.join(dir, 'known-words.json');
      fs.writeFileSync(
        cachePath,
        JSON.stringify({
          version: 1,
          words: ['知る', '猫'],
        }),
      );

      const app = createStatsApp(
        createMockTracker({
          getSessionWordsByLine: async () => [
            { lineIndex: 1, headword: '知る', occurrenceCount: 2 },
            { lineIndex: 3, headword: '猫', occurrenceCount: 1 },
            { lineIndex: 3, headword: '見る', occurrenceCount: 4 },
          ],
        }),
        { knownWordCachePath: cachePath },
      );

      const res = await app.request('/api/stats/sessions/1/known-words-timeline');
      assert.equal(res.status, 200);
      assert.deepEqual(await res.json(), [
        { linesSeen: 0, knownWordsSeen: 0, totalWordsSeen: 0 },
        { linesSeen: 1, knownWordsSeen: 2, totalWordsSeen: 2 },
        { linesSeen: 2, knownWordsSeen: 2, totalWordsSeen: 2 },
        { linesSeen: 3, knownWordsSeen: 3, totalWordsSeen: 7 },
      ]);
    });
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

  it('GET /api/stats/trends/dashboard returns chart-ready trends data', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getTrendsDashboard: async (...args: unknown[]) => {
          seenArgs = args;
          return TRENDS_DASHBOARD;
        },
      }),
    );

    const res = await app.request('/api/stats/trends/dashboard?range=90d&groupBy=month');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.deepEqual(seenArgs, ['90d', 'month']);
    assert.deepEqual(body.activity.watchTime, TRENDS_DASHBOARD.activity.watchTime);
    assert.deepEqual(body.librarySummary, TRENDS_DASHBOARD.librarySummary);
  });

  it('GET /api/stats/trends/dashboard accepts 365d range', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getTrendsDashboard: async (...args: unknown[]) => {
          seenArgs = args;
          return TRENDS_DASHBOARD;
        },
      }),
    );

    const res = await app.request('/api/stats/trends/dashboard?range=365d&groupBy=month');
    assert.equal(res.status, 200);
    assert.deepEqual(seenArgs, ['365d', 'month']);
  });

  it('GET /api/stats/trends/dashboard falls back to safe defaults for invalid params', async () => {
    let seenArgs: unknown[] = [];
    const app = createStatsApp(
      createMockTracker({
        getTrendsDashboard: async (...args: unknown[]) => {
          seenArgs = args;
          return TRENDS_DASHBOARD;
        },
      }),
    );

    const res = await app.request('/api/stats/trends/dashboard?range=weird&groupBy=year');
    assert.equal(res.status, 200);
    assert.deepEqual(seenArgs, ['30d', 'day']);
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

  it('GET /api/stats/excluded-words returns tracker exclusion rows', async () => {
    const app = createStatsApp(
      createMockTracker({
        getStatsExcludedWords: async () => [
          { headword: '猫', word: '猫', reading: 'ねこ' },
          { headword: 'する', word: 'する', reading: 'する' },
        ],
      }),
    );

    const res = await app.request('/api/stats/excluded-words');
    assert.equal(res.status, 200);
    assert.deepEqual(await res.json(), [
      { headword: '猫', word: '猫', reading: 'ねこ' },
      { headword: 'する', word: 'する', reading: 'する' },
    ]);
  });

  it('PUT /api/stats/excluded-words replaces tracker exclusion rows', async () => {
    let seenWords: unknown = null;
    const app = createStatsApp(
      createMockTracker({
        replaceStatsExcludedWords: async (words: unknown) => {
          seenWords = words;
        },
      }),
    );

    const res = await app.request('/api/stats/excluded-words', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        words: [
          { headword: '猫', word: '猫', reading: 'ねこ' },
          { headword: 'する', word: 'する', reading: 'する' },
        ],
      }),
    });

    assert.equal(res.status, 200);
    assert.deepEqual(await res.json(), { ok: true });
    assert.deepEqual(seenWords, [
      { headword: '猫', word: '猫', reading: 'ねこ' },
      { headword: 'する', word: 'する', reading: 'する' },
    ]);
  });

  it('PUT /api/stats/excluded-words rejects malformed rows', async () => {
    const app = createStatsApp(createMockTracker());

    const res = await app.request('/api/stats/excluded-words', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ words: [{ headword: '猫', word: 7, reading: 'ねこ' }] }),
    });

    assert.equal(res.status, 400);
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

  it('POST /api/stats/covers batches stored cover art without fetching missing art', async () => {
    let ensureCoverArtCalls = 0;
    const app = createStatsApp(
      createMockTracker({
        getCoverArt: async (videoId: number) =>
          videoId === 7
            ? {
                videoId,
                anilistId: null,
                coverUrl: null,
                coverBlob: Buffer.from([0x89, 0x50]),
                titleRomaji: null,
                titleEnglish: null,
                episodesTotal: null,
                fetchedAtMs: Date.now(),
              }
            : null,
        ensureCoverArt: async () => {
          ensureCoverArtCalls += 1;
          return true;
        },
      }),
    );

    const res = await app.request('/api/stats/covers', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ animeIds: [1, 99999], videoIds: [7, 99999] }),
    });

    assert.equal(res.status, 200);
    assert.deepEqual(await res.json(), {
      anime: {
        1: {
          contentType: 'image/jpeg',
          dataUrl: 'data:image/jpeg;base64,/9j/2Q==',
        },
        99999: null,
      },
      media: {
        7: {
          contentType: 'image/jpeg',
          dataUrl: 'data:image/jpeg;base64,iVA=',
        },
        99999: null,
      },
    });
    assert.equal(ensureCoverArtCalls, 0);
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

  it('POST /api/stats/mine-card falls back to Default deck for empty deck config', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          ankiConnectConfig: {
            url,
            deck: '',
            tags: ['SubMiner'],
            fields: {
              word: 'Expression',
              audio: 'ExpressionAudio',
              image: 'Picture',
              sentence: 'Sentence',
              miscInfo: 'MiscInfo',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
            isLapis: {
              enabled: true,
              sentenceCardModel: 'Lapis Morph',
            },
            isKiku: {
              enabled: true,
              fieldGrouping: 'manual',
              deleteDuplicateInAuto: true,
            },
          },
        });

        const res = await app.request('/api/stats/mine-card?mode=sentence', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            secondaryText: 'I saw a cat',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));
        assert.equal(body.noteId, 12345);

        const addNoteRequest = requests.find((request) => request.action === 'addNote');
        assert.equal(addNoteRequest?.params?.note?.deckName, 'Default');
        assert.equal(addNoteRequest?.params?.note?.modelName, 'Lapis Morph');
        assert.equal(addNoteRequest?.params?.note?.fields?.Sentence, '猫を見た');
        assert.equal(addNoteRequest?.params?.note?.fields?.IsSentenceCard, 'x');
      });
    });
  });

  it('POST /api/stats/mine-card uses Yomitan deck for direct sentence cards when config deck is empty', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          getYomitanAnkiDeckName: async () => 'Minecraft',
          ankiConnectConfig: {
            url,
            deck: '',
            tags: ['SubMiner'],
            fields: {
              word: 'Expression',
              sentence: 'Sentence',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
            isLapis: {
              enabled: true,
              sentenceCardModel: 'Lapis Morph',
            },
          },
        });

        const res = await app.request('/api/stats/mine-card?mode=sentence', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            secondaryText: 'I saw a cat',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));
        const addNoteRequest = requests.find((request) => request.action === 'addNote');
        assert.equal(addNoteRequest?.params?.note?.deckName, 'Minecraft');
      });
    });
  });

  it('POST /api/stats/mine-card resolves Anki config at request time', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          getAnkiConnectConfig: () => ({
            url,
            deck: 'Mining',
            tags: ['SubMiner'],
            fields: {
              word: 'Expression',
              sentence: 'Sentence',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
            isLapis: {
              enabled: true,
              sentenceCardModel: 'Lapis Morph',
            },
          }),
        } as Parameters<typeof createStatsApp>[1]);

        const res = await app.request('/api/stats/mine-card?mode=sentence', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            secondaryText: 'I saw a cat',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));
        const addNoteRequest = requests.find((request) => request.action === 'addNote');
        assert.equal(addNoteRequest?.params?.note?.deckName, 'Mining');
        assert.equal(addNoteRequest?.params?.note?.modelName, 'Lapis Morph');
        assert.equal(addNoteRequest?.params?.note?.fields?.SelectionText, 'I saw a cat');
      });
    });
  });

  it('POST /api/stats/mine-card adds direct sentence cards before slow media finishes', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const mediaRelease: {
          audio?: () => void;
          image?: () => void;
        } = {};
        const app = createStatsApp(createMockTracker(), {
          createMediaGenerator: () => ({
            generateAudio: async () =>
              await new Promise<Buffer>((resolve) => {
                mediaRelease.audio = () => resolve(Buffer.from('audio'));
              }),
            generateScreenshot: async () =>
              await new Promise<Buffer>((resolve) => {
                mediaRelease.image = () => resolve(Buffer.from('image'));
              }),
            generateAnimatedImage: async () => null,
          }),
          ankiConnectConfig: {
            url,
            deck: 'Mining',
            tags: ['SubMiner'],
            fields: {
              word: 'Expression',
              audio: 'ExpressionAudio',
              image: 'Picture',
              sentence: 'Sentence',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: true,
              generateImage: true,
              imageType: 'static',
            },
            isLapis: {
              enabled: true,
              sentenceCardModel: 'Lapis Morph',
            },
          },
        });

        const pendingResponse = app.request('/api/stats/mine-card?mode=sentence', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            secondaryText: 'I saw a cat',
            videoTitle: 'Episode 1',
          }),
        });

        for (let attempt = 0; attempt < 20; attempt += 1) {
          if (requests.some((request) => request.action === 'addNote')) break;
          await new Promise((resolve) => setTimeout(resolve, 1));
        }
        const addedBeforeMediaFinished = requests.some((request) => request.action === 'addNote');
        mediaRelease.audio?.();
        mediaRelease.image?.();

        const res = await pendingResponse;
        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));
        assert.equal(addedBeforeMediaFinished, true);
      });
    });
  });

  it('POST /api/stats/mine-card writes secondary subtitles to word card selection text', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          addYomitanNote: async () => 777,
          ankiConnectConfig: {
            url,
            deck: 'Mining',
            fields: {
              audio: 'ExpressionAudio',
              image: 'Picture',
              sentence: 'Sentence',
              miscInfo: 'MiscInfo',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
          },
        });

        const res = await app.request('/api/stats/mine-card?mode=word', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            secondaryText: 'I saw a cat',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));
        assert.equal(body.noteId, 777);

        const updateRequest = requests.find((request) => request.action === 'updateNoteFields');
        assert.equal(updateRequest?.params?.note?.id, 777);
        assert.equal(updateRequest?.params?.note?.fields?.Sentence, '<b>猫</b>を見た');
        assert.equal(updateRequest?.params?.note?.fields?.SelectionText, 'I saw a cat');
      });
    });
  });

  it('POST /api/stats/mine-card moves Yomitan-created word notes to the configured deck', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          addYomitanNote: async () => 777,
          ankiConnectConfig: {
            url,
            deck: 'Mining',
            fields: {
              sentence: 'Sentence',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
          },
        });

        const res = await app.request('/api/stats/mine-card?mode=word', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));

        const findCardsRequest = requests.find((request) => request.action === 'findCards');
        assert.equal(findCardsRequest?.params?.query, 'nid:777');
        const changeDeckRequest = requests.find((request) => request.action === 'changeDeck');
        assert.deepEqual(changeDeckRequest?.params?.cards, [9001]);
        assert.equal(changeDeckRequest?.params?.deck, 'Mining');
      });
    });
  });

  it('POST /api/stats/mine-card moves Yomitan-created word notes to Yomitan deck when config deck is empty', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          getYomitanAnkiDeckName: async () => 'Minecraft',
          addYomitanNote: async () => 777,
          ankiConnectConfig: {
            url,
            deck: '',
            fields: {
              sentence: 'Sentence',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
          },
        });

        const res = await app.request('/api/stats/mine-card?mode=word', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));

        const findCardsRequest = requests.find((request) => request.action === 'findCards');
        assert.equal(findCardsRequest?.params?.query, 'nid:777');
        const changeDeckRequest = requests.find((request) => request.action === 'changeDeck');
        assert.deepEqual(changeDeckRequest?.params?.cards, [9001]);
        assert.equal(changeDeckRequest?.params?.deck, 'Minecraft');
      });
    });
  });

  it('POST /api/stats/mine-card uses the full sentence as sentence-card expression', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          ankiConnectConfig: {
            url,
            deck: 'Minecraft',
            tags: ['SubMiner'],
            fields: {
              word: 'Expression',
              sentence: 'Sentence',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
            isLapis: {
              enabled: true,
              sentenceCardModel: 'Lapis Morph',
            },
          },
        });

        const res = await app.request('/api/stats/mine-card?mode=sentence', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            secondaryText: 'I saw a cat',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));

        const addNoteRequest = requests.find((request) => request.action === 'addNote');
        assert.equal(addNoteRequest?.params?.note?.deckName, 'Minecraft');
        assert.equal(addNoteRequest?.params?.note?.fields?.Expression, '猫を見た');
        assert.equal(addNoteRequest?.params?.note?.fields?.Sentence, '猫を見た');
        assert.equal(addNoteRequest?.params?.note?.fields?.SelectionText, 'I saw a cat');
        assert.equal(addNoteRequest?.params?.note?.fields?.IsSentenceCard, 'x');
      });
    });
  });

  it('POST /api/stats/mine-card fills selection text from a matching secondary sidecar subtitle', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');
      fs.writeFileSync(
        path.join(dir, 'episode.en.srt'),
        [
          '1',
          '00:00:00,800 --> 00:00:02,500',
          'I saw a cat.',
          '',
          '2',
          '00:00:03,000 --> 00:00:04,000',
          'Not this line.',
          '',
        ].join('\n'),
      );

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          ankiConnectConfig: {
            url,
            deck: 'Minecraft',
            tags: ['SubMiner'],
            fields: {
              word: 'Expression',
              sentence: 'Sentence',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
            isLapis: {
              enabled: true,
              sentenceCardModel: 'Lapis Morph',
            },
          },
          secondarySubtitleLanguages: ['en'],
        } as Parameters<typeof createStatsApp>[1]);

        const res = await app.request('/api/stats/mine-card?mode=sentence', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));

        const addNoteRequest = requests.find((request) => request.action === 'addNote');
        assert.equal(addNoteRequest?.params?.note?.deckName, 'Minecraft');
        assert.equal(addNoteRequest?.params?.note?.fields?.SelectionText, 'I saw a cat.');
      });
    });
  });

  it('POST /api/stats/mine-card does not append the next sidecar cue near a timing boundary', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');
      fs.writeFileSync(
        path.join(dir, 'episode.en.srt'),
        [
          '1',
          '00:00:00,800 --> 00:00:01,500',
          "I don't give a damn what family she's from.",
          '',
          '2',
          '00:00:01,700 --> 00:00:03,000',
          'That snobby attitude just pisses me off!',
          '',
        ].join('\n'),
      );

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          ankiConnectConfig: {
            url,
            deck: 'Minecraft',
            tags: ['SubMiner'],
            fields: {
              word: 'Expression',
              sentence: 'Sentence',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
            isLapis: {
              enabled: true,
              sentenceCardModel: 'Lapis Morph',
            },
          },
          secondarySubtitleLanguages: ['en'],
        } as Parameters<typeof createStatsApp>[1]);

        const res = await app.request('/api/stats/mine-card?mode=sentence', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '名門か何だか知らねえが',
            word: '名門',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));

        const addNoteRequest = requests.find((request) => request.action === 'addNote');
        assert.equal(
          addNoteRequest?.params?.note?.fields?.SelectionText,
          "I don't give a damn what family she's from.",
        );
      });
    });
  });

  it('POST /api/stats/mine-card writes word mining audio to SentenceAudio when present', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          addYomitanNote: async () => 777,
          createMediaGenerator: () => ({
            generateAudio: async () => Buffer.from('audio'),
            generateScreenshot: async () => null,
            generateAnimatedImage: async () => null,
          }),
          ankiConnectConfig: {
            url,
            deck: 'Mining',
            fields: {
              audio: 'ExpressionAudio',
              image: 'Picture',
              sentence: 'Sentence',
              miscInfo: 'MiscInfo',
            },
            media: {
              generateAudio: true,
              generateImage: false,
            },
          },
        });

        const res = await app.request('/api/stats/mine-card?mode=word', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));

        const updateRequest = requests.find((request) => request.action === 'updateNoteFields');
        const audioValue = updateRequest?.params?.note?.fields?.SentenceAudio;
        assert.match(audioValue ?? '', /^\[sound:subminer_audio_\d+\.mp3\]$/);
        assert.equal(updateRequest?.params?.note?.fields?.ExpressionAudio, undefined);
      });
    });
  });

  it('POST /api/stats/mine-card records timing for slow sentence mining phases', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        let now = 0;
        const timings: Array<{ mode: string; phase: string; elapsedMs: number; noteId?: number }> =
          [];
        const app = createStatsApp(createMockTracker(), {
          nowMs: () => {
            now += 10;
            return now;
          },
          onMiningTiming: (event) => {
            timings.push(event);
          },
          createMediaGenerator: () => ({
            generateAudio: async () => Buffer.from('audio'),
            generateScreenshot: async () => Buffer.from('image'),
            generateAnimatedImage: async () => Buffer.from('animated'),
          }),
          ankiConnectConfig: {
            url,
            deck: 'Mining',
            tags: ['SubMiner'],
            fields: {
              word: 'Expression',
              audio: 'ExpressionAudio',
              image: 'Picture',
              sentence: 'Sentence',
              miscInfo: 'MiscInfo',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: true,
              generateImage: true,
              imageType: 'static',
            },
            isLapis: {
              enabled: true,
              sentenceCardModel: 'Lapis Morph',
            },
          },
        });

        const res = await app.request('/api/stats/mine-card?mode=sentence', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            secondaryText: 'I saw a cat',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));
        assert.deepEqual(
          timings.map((entry) => entry.phase),
          [
            'generateAudio',
            'generateScreenshot',
            'addNote',
            'findCards',
            'changeDeck',
            'uploadAudio',
            'uploadImage',
            'updateNoteFields',
          ],
        );
        assert.ok(timings.every((entry) => entry.mode === 'sentence' && entry.elapsedMs >= 0));
        assert.equal(timings.find((entry) => entry.phase === 'addNote')?.noteId, 12345);

        const updateRequest = requests.find((request) => request.action === 'updateNoteFields');
        const audioValue = updateRequest?.params?.note?.fields?.SentenceAudio;
        assert.match(audioValue ?? '', /^\[sound:subminer_audio_\d+\.mp3\]$/);
        assert.equal(updateRequest?.params?.note?.fields?.ExpressionAudio, undefined);
      });
    });
  });

  it('POST /api/stats/mine-card only writes selection text for sentence cards', async () => {
    await withTempDir(async (dir) => {
      const sourcePath = path.join(dir, 'episode.mkv');
      fs.writeFileSync(sourcePath, 'fake media');

      await withFakeAnkiConnect(async (requests, url) => {
        const app = createStatsApp(createMockTracker(), {
          ankiConnectConfig: {
            url,
            deck: 'Mining',
            fields: {
              word: 'Expression',
              sentence: 'Sentence',
              translation: 'SelectionText',
            },
            media: {
              generateAudio: false,
              generateImage: false,
            },
            isLapis: {
              enabled: true,
              sentenceCardModel: 'Lapis Morph',
            },
          },
        });

        const res = await app.request('/api/stats/mine-card?mode=audio', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sourcePath,
            startMs: 1_000,
            endMs: 2_000,
            sentence: '猫を見た',
            word: '猫',
            secondaryText: 'I saw a cat',
            videoTitle: 'Episode 1',
          }),
        });

        const body = await res.json();
        assert.equal(res.status, 200, JSON.stringify(body));

        const addNoteRequest = requests.find((request) => request.action === 'addNote');
        assert.equal(addNoteRequest?.params?.note?.fields?.SelectionText, undefined);
        assert.equal(addNoteRequest?.params?.note?.fields?.IsAudioCard, 'x');
      });
    });
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

  it('GET /api/stats/anilist/search uses the configured AniList rate limiter', async () => {
    const originalFetch = globalThis.fetch;
    let acquireCalls = 0;
    let recordCalls = 0;
    globalThis.fetch = (async () =>
      new Response(
        JSON.stringify({
          data: {
            Page: {
              media: [{ id: 21858, title: { romaji: 'Little Witch Academia' } }],
            },
          },
        }),
        {
          status: 200,
          headers: { 'Content-Type': 'application/json', 'X-RateLimit-Remaining': '29' },
        },
      )) as typeof fetch;

    try {
      const app = createStatsApp(createMockTracker(), {
        anilistRateLimiter: {
          acquire: async () => {
            acquireCalls += 1;
          },
          recordResponse: () => {
            recordCalls += 1;
          },
        },
      });
      const res = await app.request('/api/stats/anilist/search?q=Little%20Witch%20Academia');

      assert.equal(res.status, 200);
      assert.equal(acquireCalls, 1);
      assert.equal(recordCalls, 1);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('POST /api/stats/anki/notesInfo resolves stale note ids through the configured alias resolver', async () => {
    const originalFetch = globalThis.fetch;
    const requests: unknown[] = [];
    globalThis.fetch = (async (_input: RequestInfo | URL, init?: RequestInit) => {
      requests.push(init?.body ? JSON.parse(String(init.body)) : null);
      return new Response(
        JSON.stringify({
          result: [
            {
              noteId: 222,
              fields: {
                Expression: { value: '呪い' },
              },
            },
          ],
        }),
        {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        },
      );
    }) as typeof fetch;

    try {
      const app = createStatsApp(createMockTracker(), {
        resolveAnkiNoteId: (noteId) => (noteId === 111 ? 222 : noteId),
      });
      const res = await app.request('/api/stats/anki/notesInfo', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ noteIds: [111] }),
      });

      assert.equal(res.status, 200);
      assert.deepEqual(requests, [
        {
          action: 'notesInfo',
          version: 6,
          params: { notes: [222] },
        },
      ]);
      assert.deepEqual(await res.json(), [
        {
          noteId: 222,
          fields: {
            Expression: { value: '呪い' },
          },
          preview: {
            word: '呪い',
            sentence: '',
            translation: '',
          },
        },
      ]);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('POST /api/stats/anki/notesInfo returns preview fields using configured word and sentence field names', async () => {
    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async () =>
      new Response(
        JSON.stringify({
          result: [
            {
              noteId: 333,
              fields: {
                TargetWord: { value: '<span>連れる</span>' },
                Quote: { value: '<div>このまま<b>連れてって</b></div>' },
                SelectionText: { value: 'to take along' },
              },
            },
          ],
        }),
        {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        },
      )) as typeof fetch;

    try {
      const app = createStatsApp(createMockTracker(), {
        ankiConnectConfig: {
          fields: {
            word: 'TargetWord',
            sentence: 'Quote',
            translation: 'SelectionText',
          },
        },
      });
      const res = await app.request('/api/stats/anki/notesInfo', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ noteIds: [333] }),
      });

      assert.equal(res.status, 200);
      assert.deepEqual(await res.json(), [
        {
          noteId: 333,
          fields: {
            TargetWord: { value: '<span>連れる</span>' },
            Quote: { value: '<div>このまま<b>連れてって</b></div>' },
            SelectionText: { value: 'to take along' },
          },
          preview: {
            word: '連れる',
            sentence: 'このまま 連れてって',
            translation: 'to take along',
          },
        },
      ]);
    } finally {
      globalThis.fetch = originalFetch;
    }
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

  it('starts the stats server with Bun.serve', () => {
    type BunRuntime = {
      Bun: {
        serve: (options: { fetch: unknown; port: number; hostname: string }) => {
          stop: () => void;
        };
      };
    };

    const bun = globalThis as typeof globalThis & BunRuntime;
    const originalServe = bun.Bun.serve;
    let servedWith: { fetch: unknown; port: number; hostname: string } | null = null;
    let stopCalls = 0;

    bun.Bun.serve = (options: { fetch: unknown; port: number; hostname: string }) => {
      servedWith = options;
      return {
        stop: () => {
          stopCalls += 1;
        },
      };
    };

    try {
      const server = startStatsServer({
        port: 3210,
        staticDir: fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-stats-server-start-')),
        tracker: createMockTracker(),
      });

      if (servedWith === null) {
        throw new Error('expected Bun.serve to be called');
      }

      const servedOptions = servedWith as {
        fetch: unknown;
        port: number;
        hostname: string;
      };
      assert.equal(servedOptions.port, 3210);
      assert.equal(servedOptions.hostname, '127.0.0.1');
      assert.equal(typeof servedOptions.fetch, 'function');

      server.close();
      assert.equal(stopCalls, 1);
    } finally {
      bun.Bun.serve = originalServe;
    }
  });

  it('falls back to node:http when Bun.serve is unavailable', () => {
    type BunRuntime = {
      Bun: {
        serve?: (options: { fetch: unknown; port: number; hostname: string }) => {
          stop: () => void;
        };
      };
    };

    const bun = globalThis as typeof globalThis & BunRuntime;
    const originalServe = bun.Bun.serve;
    const originalCreateServer = http.createServer;
    let listenedWith: { port: number; hostname: string } | null = null;
    let closeCalls = 0;
    bun.Bun.serve = undefined;
    (
      http as typeof http & {
        createServer: typeof http.createServer;
      }
    ).createServer = (() =>
      ({
        listen: (port: number, hostname: string) => {
          listenedWith = { port, hostname };
        },
        close: () => {
          closeCalls += 1;
        },
      }) as unknown as ReturnType<typeof http.createServer>) as typeof http.createServer;

    try {
      const server = startStatsServer({
        port: 0,
        staticDir: fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-stats-server-node-')),
        tracker: createMockTracker(),
      });

      assert.deepEqual(listenedWith, { port: 0, hostname: '127.0.0.1' });
      server.close();
      assert.equal(closeCalls, 1);
    } finally {
      bun.Bun.serve = originalServe;
      (
        http as typeof http & {
          createServer: typeof http.createServer;
        }
      ).createServer = originalCreateServer;
    }
  });
});
