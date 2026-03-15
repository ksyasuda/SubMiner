# Stats Dashboard Redesign — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Restructure the stats dashboard from word-heavy metrics to an anime-centric 5-tab layout (Overview, Anime, Trends, Vocabulary, Sessions).

**Architecture:** The backend already has anime-level queries (`getAnimeLibrary`, `getAnimeDetail`, `getAnimeEpisodes`) and occurrence queries (`getWordOccurrences`, `getKanjiOccurrences`) but no API endpoints for anime. The frontend needs new types, hooks, components, and data builders. Work proceeds bottom-up: backend API → frontend types → hooks → data builders → components.

**Tech Stack:** Hono (backend API), React + Recharts + Tailwind/Catppuccin (frontend), node:test (backend tests), Vitest (frontend tests)

---

## Task 1: Add Anime API Endpoints to Stats Server

**Files:**
- Modify: `src/core/services/stats-server.ts`
- Modify: `src/core/services/__tests__/stats-server.test.ts`

**Context:** `query.ts` already exports `getAnimeLibrary()`, `getAnimeDetail(db, animeId)`, `getAnimeEpisodes(db, animeId)`. The stats server needs endpoints to expose them. Also need an anime cover art endpoint — anime links to videos, and videos link to `imm_media_art`.

**Step 1: Write failing tests for the 3 new anime endpoints**

Add to `src/core/services/__tests__/stats-server.test.ts`:

```typescript
test('GET /api/stats/anime returns anime library', async () => {
  const res = await app.request('/api/stats/anime');
  assert.equal(res.status, 200);
  const body = await res.json();
  assert.ok(Array.isArray(body));
});

test('GET /api/stats/anime/:animeId returns anime detail', async () => {
  // Use animeId from seeded data
  const res = await app.request('/api/stats/anime/1');
  assert.equal(res.status, 200);
  const body = await res.json();
  assert.ok(body.detail);
  assert.ok(Array.isArray(body.episodes));
});

test('GET /api/stats/anime/:animeId returns 404 for missing anime', async () => {
  const res = await app.request('/api/stats/anime/99999');
  assert.equal(res.status, 404);
});
```

**Step 2: Run tests to verify they fail**

Run: `node --test src/core/services/__tests__/stats-server.test.ts`
Expected: FAIL — 404 for all anime endpoints

**Step 3: Add the anime endpoints to stats-server.ts**

Add after the existing media endpoints (around line 165):

```typescript
app.get('/api/stats/anime', async (c) => {
  const rows = getAnimeLibrary(tracker.db);
  return c.json(rows);
});

app.get('/api/stats/anime/:animeId', async (c) => {
  const animeId = parseIntQuery(c.req.param('animeId'), 0);
  if (animeId <= 0) return c.body(null, 400);
  const detail = getAnimeDetail(tracker.db, animeId);
  if (!detail) return c.body(null, 404);
  const episodes = getAnimeEpisodes(tracker.db, animeId);
  return c.json({ detail, episodes });
});

app.get('/api/stats/anime/:animeId/cover', async (c) => {
  const animeId = parseIntQuery(c.req.param('animeId'), 0);
  if (animeId <= 0) return c.body(null, 404);
  const art = getAnimeCoverArt(tracker.db, animeId);
  if (!art?.coverBlob) return c.body(null, 404);
  return new Response(new Uint8Array(art.coverBlob), {
    headers: {
      'Content-Type': 'image/jpeg',
      'Cache-Control': 'public, max-age=86400',
    },
  });
});
```

Note: `getAnimeCoverArt` may need to be added to `query.ts` — it should look up the first video for the anime and return its cover art. Check if this already exists; if not, add it:

```typescript
export function getAnimeCoverArt(db: DatabaseSync, animeId: number): MediaArtRow | null {
  return db.prepare(`
    SELECT a.video_id, a.anilist_id, a.cover_url, a.cover_blob,
           a.title_romaji, a.title_english, a.episodes_total, a.fetched_at_ms
    FROM imm_media_art a
    JOIN imm_videos v ON v.video_id = a.video_id
    WHERE v.anime_id = ?
    AND a.cover_blob IS NOT NULL
    LIMIT 1
  `).get(animeId) as MediaArtRow | null;
}
```

Import the new query functions at the top of `stats-server.ts`.

**Step 4: Run tests to verify they pass**

Run: `node --test src/core/services/__tests__/stats-server.test.ts`
Expected: All tests PASS

**Step 5: Commit**

```bash
git add src/core/services/stats-server.ts src/core/services/__tests__/stats-server.test.ts src/core/services/immersion-tracker/query.ts
git commit -m "feat(stats): add anime API endpoints to stats server"
```

---

## Task 2: Add Anime Words and Anime Rollups Query + Endpoints

**Files:**
- Modify: `src/core/services/immersion-tracker/query.ts`
- Modify: `src/core/services/immersion-tracker/types.ts`
- Modify: `src/core/services/stats-server.ts`
- Modify: `src/core/services/__tests__/stats-server.test.ts`

**Context:** The anime detail view needs "top words from this anime" and daily rollups scoped to an anime. Word occurrences already join through `imm_word_line_occurrences` → `imm_subtitle_lines` → `anime_id`.

**Step 1: Add query functions to query.ts**

```typescript
export interface AnimeWordRow {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  partOfSpeech: string | null;
  frequency: number;
}

export function getAnimeWords(db: DatabaseSync, animeId: number, limit = 50): AnimeWordRow[] {
  return db.prepare(`
    SELECT w.id AS wordId, w.headword, w.word, w.reading, w.part_of_speech AS partOfSpeech,
           SUM(o.occurrence_count) AS frequency
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_words w ON w.id = o.word_id
    WHERE sl.anime_id = ?
    GROUP BY w.id
    ORDER BY frequency DESC
    LIMIT ?
  `).all(animeId, limit) as AnimeWordRow[];
}

export function getAnimeDailyRollups(db: DatabaseSync, animeId: number, limit = 90): ImmersionSessionRollupRow[] {
  return db.prepare(`
    SELECT r.rollup_day AS rollupDayOrMonth, r.video_id AS videoId,
           r.total_sessions AS totalSessions, r.total_active_min AS totalActiveMin,
           r.total_lines_seen AS totalLinesSeen, r.total_words_seen AS totalWordsSeen,
           r.total_tokens_seen AS totalTokensSeen, r.total_cards AS totalCards,
           r.cards_per_hour AS cardsPerHour, r.words_per_min AS wordsPerMin,
           r.lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups r
    JOIN imm_videos v ON v.video_id = r.video_id
    WHERE v.anime_id = ?
    ORDER BY r.rollup_day DESC
    LIMIT ?
  `).all(animeId, limit) as ImmersionSessionRollupRow[];
}
```

Add `AnimeWordRow` to the exports in `types.ts` if needed.

**Step 2: Add API endpoints**

In `stats-server.ts`, add:

```typescript
app.get('/api/stats/anime/:animeId/words', async (c) => {
  const animeId = parseIntQuery(c.req.param('animeId'), 0);
  const limit = parseIntQuery(c.req.query('limit'), 50, 200);
  if (animeId <= 0) return c.body(null, 400);
  return c.json(getAnimeWords(tracker.db, animeId, limit));
});

app.get('/api/stats/anime/:animeId/rollups', async (c) => {
  const animeId = parseIntQuery(c.req.param('animeId'), 0);
  const limit = parseIntQuery(c.req.query('limit'), 90, 365);
  if (animeId <= 0) return c.body(null, 400);
  return c.json(getAnimeDailyRollups(tracker.db, animeId, limit));
});
```

**Step 3: Write tests and verify**

Run: `node --test src/core/services/__tests__/stats-server.test.ts`

**Step 4: Commit**

```bash
git add src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker/types.ts src/core/services/stats-server.ts src/core/services/__tests__/stats-server.test.ts
git commit -m "feat(stats): add anime words and rollups query + endpoints"
```

---

## Task 3: Extend Overview Endpoint with Episodes and Active Anime

**Files:**
- Modify: `src/core/services/immersion-tracker/query.ts`
- Modify: `src/core/services/stats-server.ts`
- Modify: `src/core/services/__tests__/stats-server.test.ts`

**Context:** The overview endpoint currently returns `{ sessions, rollups, hints }`. We need to add `episodesToday` and `activeAnimeCount` to the hints.

**Step 1: Extend `getQueryHints` in query.ts**

The current `getQueryHints` returns `{ totalSessions, activeSessions }`. Add:

```typescript
// Episodes today: count distinct video_ids with sessions started today
const today = Math.floor(Date.now() / 86400000);
const episodesToday = (db.prepare(`
  SELECT COUNT(DISTINCT v.video_id) AS count
  FROM imm_sessions s
  JOIN imm_videos v ON v.video_id = s.video_id
  WHERE CAST(s.started_at_ms / 86400000 AS INTEGER) = ?
`).get(today) as { count: number })?.count ?? 0;

// Active anime: anime with sessions in last 30 days
const thirtyDaysAgoMs = Date.now() - 30 * 86400000;
const activeAnimeCount = (db.prepare(`
  SELECT COUNT(DISTINCT v.anime_id) AS count
  FROM imm_sessions s
  JOIN imm_videos v ON v.video_id = s.video_id
  WHERE v.anime_id IS NOT NULL
  AND s.started_at_ms >= ?
`).get(thirtyDaysAgoMs) as { count: number })?.count ?? 0;
```

Return these as part of the hints object.

**Step 2: Write test, run, verify**

Run: `node --test src/core/services/__tests__/stats-server.test.ts`

**Step 3: Commit**

```bash
git add src/core/services/immersion-tracker/query.ts src/core/services/stats-server.ts src/core/services/__tests__/stats-server.test.ts
git commit -m "feat(stats): add episodes today and active anime to overview hints"
```

---

## Task 4: Extend Vocabulary Endpoint with POS Data

**Files:**
- Modify: `src/core/services/immersion-tracker/query.ts` (function `getVocabularyStats`)
- Modify: `src/core/services/immersion-tracker/types.ts` (`VocabularyStatsRow`)
- Modify: `src/core/services/__tests__/stats-server.test.ts`

**Context:** The vocabulary endpoint currently returns headword, word, reading, frequency, firstSeen, lastSeen. It needs to also return `partOfSpeech`, `pos1`, `pos2`, `pos3` and support a `?excludePos=particle` query param for filtering.

**Step 1: Update `VocabularyStatsRow` in types.ts**

Add fields: `partOfSpeech: string | null`, `pos1: string | null`, `pos2: string | null`, `pos3: string | null`.

**Step 2: Update `getVocabularyStats` query in query.ts**

Add `part_of_speech AS partOfSpeech, pos1, pos2, pos3` to the SELECT. Add optional POS filtering parameter.

**Step 3: Update the API endpoint in stats-server.ts**

Pass the `excludePos` query param through to the query function.

**Step 4: Update frontend type `VocabularyEntry` in stats/src/types/stats.ts**

Add: `partOfSpeech: string | null`, `pos1: string | null`, `pos2: string | null`, `pos3: string | null`.

**Step 5: Test and commit**

```bash
git commit -m "feat(stats): add POS data to vocabulary endpoint and support filtering"
```

---

## Task 5: Add Streak Calendar Query + Endpoint

**Files:**
- Modify: `src/core/services/immersion-tracker/query.ts`
- Modify: `src/core/services/immersion-tracker/types.ts`
- Modify: `src/core/services/stats-server.ts`
- Modify: `src/core/services/__tests__/stats-server.test.ts`

**Context:** The streak calendar needs a map of `{ epochDay → totalActiveMin }` for the last 90 days.

**Step 1: Add query function**

```typescript
export interface StreakCalendarRow {
  epochDay: number;
  totalActiveMin: number;
}

export function getStreakCalendar(db: DatabaseSync, days = 90): StreakCalendarRow[] {
  const cutoffDay = Math.floor(Date.now() / 86400000) - days;
  return db.prepare(`
    SELECT rollup_day AS epochDay, SUM(total_active_min) AS totalActiveMin
    FROM imm_daily_rollups
    WHERE rollup_day >= ?
    GROUP BY rollup_day
    ORDER BY rollup_day ASC
  `).all(cutoffDay) as StreakCalendarRow[];
}
```

**Step 2: Add endpoint**

```typescript
app.get('/api/stats/streak-calendar', async (c) => {
  const days = parseIntQuery(c.req.query('days'), 90, 365);
  return c.json(getStreakCalendar(tracker.db, days));
});
```

**Step 3: Test and commit**

```bash
git commit -m "feat(stats): add streak calendar endpoint"
```

---

## Task 6: Add Trend Episode/Anime Queries + Endpoints

**Files:**
- Modify: `src/core/services/immersion-tracker/query.ts`
- Modify: `src/core/services/stats-server.ts`
- Modify: `src/core/services/__tests__/stats-server.test.ts`

**Context:** Trends tab needs new data: episodes per day, new anime started per day, watch time per anime (stacked).

**Step 1: Add query functions**

```typescript
export interface EpisodesPerDayRow {
  epochDay: number;
  episodeCount: number;
}

export function getEpisodesPerDay(db: DatabaseSync, limit = 90): EpisodesPerDayRow[] {
  return db.prepare(`
    SELECT CAST(s.started_at_ms / 86400000 AS INTEGER) AS epochDay,
           COUNT(DISTINCT s.video_id) AS episodeCount
    FROM imm_sessions s
    GROUP BY epochDay
    ORDER BY epochDay DESC
    LIMIT ?
  `).all(limit) as EpisodesPerDayRow[];
}

export interface NewAnimePerDayRow {
  epochDay: number;
  newAnimeCount: number;
}

export function getNewAnimePerDay(db: DatabaseSync, limit = 90): NewAnimePerDayRow[] {
  return db.prepare(`
    SELECT CAST(MIN(s.started_at_ms) / 86400000 AS INTEGER) AS epochDay,
           COUNT(*) AS newAnimeCount
    FROM (
      SELECT v.anime_id, MIN(s.started_at_ms) AS started_at_ms
      FROM imm_sessions s
      JOIN imm_videos v ON v.video_id = s.video_id
      WHERE v.anime_id IS NOT NULL
      GROUP BY v.anime_id
    ) s
    GROUP BY epochDay
    ORDER BY epochDay DESC
    LIMIT ?
  `).all(limit) as NewAnimePerDayRow[];
}

export interface WatchTimePerAnimeRow {
  epochDay: number;
  animeId: number;
  animeTitle: string;
  totalActiveMin: number;
}

export function getWatchTimePerAnime(db: DatabaseSync, limit = 90): WatchTimePerAnimeRow[] {
  const cutoffDay = Math.floor(Date.now() / 86400000) - limit;
  return db.prepare(`
    SELECT r.rollup_day AS epochDay, a.anime_id AS animeId,
           a.canonical_title AS animeTitle,
           SUM(r.total_active_min) AS totalActiveMin
    FROM imm_daily_rollups r
    JOIN imm_videos v ON v.video_id = r.video_id
    JOIN imm_anime a ON a.anime_id = v.anime_id
    WHERE r.rollup_day >= ?
    GROUP BY r.rollup_day, a.anime_id
    ORDER BY r.rollup_day ASC
  `).all(cutoffDay) as WatchTimePerAnimeRow[];
}
```

**Step 2: Add endpoints**

```typescript
app.get('/api/stats/trends/episodes-per-day', async (c) => {
  const limit = parseIntQuery(c.req.query('limit'), 90, 365);
  return c.json(getEpisodesPerDay(tracker.db, limit));
});

app.get('/api/stats/trends/new-anime-per-day', async (c) => {
  const limit = parseIntQuery(c.req.query('limit'), 90, 365);
  return c.json(getNewAnimePerDay(tracker.db, limit));
});

app.get('/api/stats/trends/watch-time-per-anime', async (c) => {
  const limit = parseIntQuery(c.req.query('limit'), 90, 365);
  return c.json(getWatchTimePerAnime(tracker.db, limit));
});
```

**Step 3: Test and commit**

```bash
git commit -m "feat(stats): add episode and anime trend query endpoints"
```

---

## Task 7: Add Word/Kanji Detail Queries + Endpoints

**Files:**
- Modify: `src/core/services/immersion-tracker/query.ts`
- Modify: `src/core/services/stats-server.ts`
- Modify: `src/core/services/__tests__/stats-server.test.ts`

**Context:** Clicking a word in the vocab tab should show full detail: POS, frequency, anime appearances, example lines, similar words.

**Step 1: Add query functions**

```typescript
export interface WordDetailRow {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  partOfSpeech: string | null;
  pos1: string | null;
  pos2: string | null;
  pos3: string | null;
  frequency: number;
  firstSeen: number;
  lastSeen: number;
}

export function getWordDetail(db: DatabaseSync, wordId: number): WordDetailRow | null {
  return db.prepare(`
    SELECT id AS wordId, headword, word, reading,
           part_of_speech AS partOfSpeech, pos1, pos2, pos3,
           frequency, first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_words WHERE id = ?
  `).get(wordId) as WordDetailRow | null;
}

export interface WordAnimeAppearanceRow {
  animeId: number;
  animeTitle: string;
  occurrenceCount: number;
}

export function getWordAnimeAppearances(db: DatabaseSync, wordId: number): WordAnimeAppearanceRow[] {
  return db.prepare(`
    SELECT a.anime_id AS animeId, a.canonical_title AS animeTitle,
           SUM(o.occurrence_count) AS occurrenceCount
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_anime a ON a.anime_id = sl.anime_id
    WHERE o.word_id = ?
    GROUP BY a.anime_id
    ORDER BY occurrenceCount DESC
  `).all(wordId) as WordAnimeAppearanceRow[];
}

export interface SimilarWordRow {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  frequency: number;
}

export function getSimilarWords(db: DatabaseSync, wordId: number, limit = 10): SimilarWordRow[] {
  const word = db.prepare('SELECT headword, reading FROM imm_words WHERE id = ?').get(wordId) as { headword: string; reading: string } | null;
  if (!word) return [];
  // Words sharing kanji characters or same reading
  return db.prepare(`
    SELECT id AS wordId, headword, word, reading, frequency
    FROM imm_words
    WHERE id != ?
    AND (reading = ? OR headword LIKE ? OR headword LIKE ?)
    ORDER BY frequency DESC
    LIMIT ?
  `).all(
    wordId,
    word.reading,
    `%${word.headword.charAt(0)}%`,
    `%${word.headword.charAt(word.headword.length - 1)}%`,
    limit
  ) as SimilarWordRow[];
}
```

Add analogous `getKanjiDetail`, `getKanjiAnimeAppearances`, `getKanjiWords` functions.

**Step 2: Add endpoints**

```typescript
app.get('/api/stats/vocabulary/:wordId/detail', async (c) => {
  const wordId = parseIntQuery(c.req.param('wordId'), 0);
  if (wordId <= 0) return c.body(null, 400);
  const detail = getWordDetail(tracker.db, wordId);
  if (!detail) return c.body(null, 404);
  const animeAppearances = getWordAnimeAppearances(tracker.db, wordId);
  const similarWords = getSimilarWords(tracker.db, wordId);
  return c.json({ detail, animeAppearances, similarWords });
});

app.get('/api/stats/kanji/:kanjiId/detail', async (c) => {
  const kanjiId = parseIntQuery(c.req.param('kanjiId'), 0);
  if (kanjiId <= 0) return c.body(null, 400);
  const detail = getKanjiDetail(tracker.db, kanjiId);
  if (!detail) return c.body(null, 404);
  const animeAppearances = getKanjiAnimeAppearances(tracker.db, kanjiId);
  const words = getKanjiWords(tracker.db, kanjiId);
  return c.json({ detail, animeAppearances, words });
});
```

**Step 3: Test and commit**

```bash
git commit -m "feat(stats): add word and kanji detail endpoints"
```

---

## Task 8: Update Frontend Types

**Files:**
- Modify: `stats/src/types/stats.ts`

**Context:** Add all new types needed by the frontend for the redesigned dashboard.

**Step 1: Add new types**

```typescript
// Anime types
export interface AnimeLibraryItem {
  animeId: number;
  canonicalTitle: string;
  anilistId: number | null;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  episodeCount: number;
  lastWatchedMs: number;
}

export interface AnimeDetailData {
  detail: {
    animeId: number;
    canonicalTitle: string;
    anilistId: number | null;
    titleRomaji: string | null;
    titleEnglish: string | null;
    titleNative: string | null;
    totalSessions: number;
    totalActiveMs: number;
    totalCards: number;
    totalWordsSeen: number;
    totalLinesSeen: number;
    totalLookupCount: number;
    totalLookupHits: number;
    episodeCount: number;
    lastWatchedMs: number;
  };
  episodes: AnimeEpisode[];
}

export interface AnimeEpisode {
  videoId: number;
  parsedEpisode: number | null;
  parsedSeason: number | null;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  lastWatchedMs: number;
}

export interface AnimeWord {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  partOfSpeech: string | null;
  frequency: number;
}

// Streak calendar
export interface StreakCalendarDay {
  epochDay: number;
  totalActiveMin: number;
}

// Trend types
export interface EpisodesPerDay {
  epochDay: number;
  episodeCount: number;
}

export interface NewAnimePerDay {
  epochDay: number;
  newAnimeCount: number;
}

export interface WatchTimePerAnime {
  epochDay: number;
  animeId: number;
  animeTitle: string;
  totalActiveMin: number;
}

// Word/Kanji detail
export interface WordDetailData {
  detail: {
    wordId: number;
    headword: string;
    word: string;
    reading: string;
    partOfSpeech: string | null;
    pos1: string | null;
    pos2: string | null;
    pos3: string | null;
    frequency: number;
    firstSeen: number;
    lastSeen: number;
  };
  animeAppearances: Array<{
    animeId: number;
    animeTitle: string;
    occurrenceCount: number;
  }>;
  similarWords: Array<{
    wordId: number;
    headword: string;
    word: string;
    reading: string;
    frequency: number;
  }>;
}
```

Update `VocabularyEntry` to include POS fields.

**Step 2: Commit**

```bash
git commit -m "feat(stats): add frontend types for anime-centric dashboard"
```

---

## Task 9: Update API Client and IPC Client

**Files:**
- Modify: `stats/src/lib/api-client.ts`
- Modify: `stats/src/lib/ipc-client.ts`

**Context:** Both clients need methods for the new endpoints.

**Step 1: Add new methods to api-client.ts**

```typescript
getAnimeLibrary: () => fetchJson<AnimeLibraryItem[]>('/api/stats/anime'),
getAnimeDetail: (animeId: number) => fetchJson<AnimeDetailData>(`/api/stats/anime/${animeId}`),
getAnimeWords: (animeId: number, limit = 50) => fetchJson<AnimeWord[]>(`/api/stats/anime/${animeId}/words?limit=${limit}`),
getAnimeRollups: (animeId: number, limit = 90) => fetchJson<DailyRollup[]>(`/api/stats/anime/${animeId}/rollups?limit=${limit}`),
getAnimeCover: (animeId: number) => `/api/stats/anime/${animeId}/cover`,
getStreakCalendar: (days = 90) => fetchJson<StreakCalendarDay[]>(`/api/stats/streak-calendar?days=${days}`),
getEpisodesPerDay: (limit = 90) => fetchJson<EpisodesPerDay[]>(`/api/stats/trends/episodes-per-day?limit=${limit}`),
getNewAnimePerDay: (limit = 90) => fetchJson<NewAnimePerDay[]>(`/api/stats/trends/new-anime-per-day?limit=${limit}`),
getWatchTimePerAnime: (limit = 90) => fetchJson<WatchTimePerAnime[]>(`/api/stats/trends/watch-time-per-anime?limit=${limit}`),
getWordDetail: (wordId: number) => fetchJson<WordDetailData>(`/api/stats/vocabulary/${wordId}/detail`),
getKanjiDetail: (kanjiId: number) => fetchJson<KanjiDetailData>(`/api/stats/kanji/${kanjiId}/detail`),
```

Mirror the same methods in `ipc-client.ts`.

**Step 2: Commit**

```bash
git commit -m "feat(stats): add anime and detail methods to API clients"
```

---

## Task 10: Add Frontend Hooks

**Files:**
- Create: `stats/src/hooks/useAnimeLibrary.ts`
- Create: `stats/src/hooks/useAnimeDetail.ts`
- Create: `stats/src/hooks/useStreakCalendar.ts`
- Create: `stats/src/hooks/useWordDetail.ts`
- Create: `stats/src/hooks/useKanjiDetail.ts`

**Context:** Follow the same pattern as existing hooks (e.g., `useMediaLibrary`, `useMediaDetail`). Each hook: fetches on mount or param change, returns `{ data, loading, error }`, handles cleanup.

**Step 1: Create hooks**

Pattern for each (example `useAnimeLibrary`):

```typescript
import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { AnimeLibraryItem } from '../types/stats';

export function useAnimeLibrary() {
  const [anime, setAnime] = useState<AnimeLibraryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    const client = getStatsClient();
    client.getAnimeLibrary()
      .then((data) => { if (!cancelled) setAnime(data); })
      .catch((err) => { if (!cancelled) setError(String(err)); })
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, []);

  return { anime, loading, error };
}
```

Follow same pattern for `useAnimeDetail(animeId)`, `useStreakCalendar()`, `useWordDetail(wordId)`, `useKanjiDetail(kanjiId)`.

**Step 2: Commit**

```bash
git commit -m "feat(stats): add frontend hooks for anime, streak, and word detail"
```

---

## Task 11: Update Dashboard Data Builders

**Files:**
- Modify: `stats/src/lib/dashboard-data.ts`
- Modify: `stats/src/lib/dashboard-data.test.ts`

**Context:** Update `buildOverviewSummary` to include episodes today and active anime. Add builder for streak calendar data. Extend `buildTrendDashboard` for new chart series.

**Step 1: Update OverviewSummary interface**

Add fields: `episodesToday: number`, `activeAnimeCount: number`. Remove `todayWords`, `weekWords`. Remove `recentWords` chart data.

**Step 2: Update buildOverviewSummary**

Pull `episodesToday` and `activeAnimeCount` from `data.hints`.

**Step 3: Add streak calendar builder**

```typescript
export interface StreakCalendarPoint {
  date: string;  // YYYY-MM-DD
  value: number; // active minutes
}

export function buildStreakCalendar(days: StreakCalendarDay[]): StreakCalendarPoint[] {
  return days.map(d => ({
    date: epochDayToDate(d.epochDay).toISOString().slice(0, 10),
    value: d.totalActiveMin,
  }));
}
```

**Step 4: Update tests, run, verify**

Run: `npx vitest run stats/src/lib/dashboard-data.test.ts`

**Step 5: Commit**

```bash
git commit -m "feat(stats): update dashboard data builders for anime-centric overview"
```

---

## Task 12: Update Tab Bar and App Router

**Files:**
- Modify: `stats/src/App.tsx`
- Modify: `stats/src/components/layout/TabBar.tsx`

**Context:** Change tabs from `['overview', 'library', 'trends', 'vocabulary']` to `['overview', 'anime', 'trends', 'vocabulary', 'sessions']`.

**Step 1: Update TabBar**

Change `TabId` type to `'overview' | 'anime' | 'trends' | 'vocabulary' | 'sessions'`. Update tab labels.

**Step 2: Update App.tsx**

Replace `LibraryTab` import with `AnimeTab` (to be created). Add `SessionsTab` import. Update conditional rendering.

Note: `AnimeTab` doesn't exist yet — create a placeholder that renders "Anime tab coming soon" for now. Wire up `SessionsTab`.

**Step 3: Commit**

```bash
git commit -m "feat(stats): update tab bar to 5-tab anime-centric layout"
```

---

## Task 13: Redesign Overview Tab

**Files:**
- Modify: `stats/src/components/overview/OverviewTab.tsx`
- Modify: `stats/src/components/overview/HeroStats.tsx`
- Create: `stats/src/components/overview/StreakCalendar.tsx`

**Context:** Update hero stats to show the 6 new metrics. Add streak calendar. Remove words chart. Keep watch time chart and recent sessions.

**Step 1: Update HeroStats**

Change the 6 cards to:
1. Watch Time Today
2. Cards Mined Today
3. Sessions Today
4. Episodes Today
5. Current Streak
6. Active Anime

**Step 2: Create StreakCalendar component**

GitHub-contributions-style heatmap. 90 days, 7 rows (days of week), colored by intensity (Catppuccin palette: ctp-surface0 for empty, ctp-green shades for activity levels).

Use `useStreakCalendar()` hook to fetch data.

**Step 3: Update OverviewTab layout**

- HeroStats (6 cards)
- Watch Time Chart (keep)
- Streak Calendar (new)
- Tracking Snapshot (updated: total sessions, total episodes, all-time hours, active days, total cards)
- Recent Sessions (keep, add episode number to display)

**Step 4: Commit**

```bash
git commit -m "feat(stats): redesign overview tab with episodes, streak calendar"
```

---

## Task 14: Build Anime Tab

**Files:**
- Create: `stats/src/components/anime/AnimeTab.tsx`
- Create: `stats/src/components/anime/AnimeCard.tsx`
- Create: `stats/src/components/anime/AnimeDetailView.tsx`
- Create: `stats/src/components/anime/AnimeHeader.tsx`
- Create: `stats/src/components/anime/EpisodeList.tsx`
- Create: `stats/src/components/anime/AnimeWordList.tsx`
- Create: `stats/src/components/anime/AnimeCoverImage.tsx`

**Context:** This replaces the Library tab. Reuse patterns from `LibraryTab` / `MediaDetailView` but centered on `anime_id` instead of `video_id`.

**Step 1: AnimeTab (grid view)**

- Search input for filtering by title
- Sort dropdown: last watched, watch time, cards, progress
- Responsive grid of AnimeCard components
- Total count + total watch time header

**Step 2: AnimeCard**

- Cover art via `AnimeCoverImage` (fetches from `/api/stats/anime/:animeId/cover`)
- Title, progress bar (`episodeCount / episodesTotal`), watch time, cards mined
- Click handler to enter detail view

**Step 3: AnimeDetailView**

- AnimeHeader: cover art, titles (romaji/english/native), AniList link, progress bar
- Stats row: 6 StatCards (watch time, cards, words, lookup rate, sessions, avg session)
- EpisodeList: table of episodes from `AnimeDetailData.episodes`
- Watch time chart using anime rollups
- AnimeWordList: top words from `getAnimeWords`, each clickable (opens vocab detail)
- Mining efficiency chart: cards per hour / cards per episode

**Step 4: Commit**

```bash
git commit -m "feat(stats): build anime tab with grid, detail, episodes, words"
```

---

## Task 15: Extend Trends Tab with New Charts

**Files:**
- Modify: `stats/src/components/trends/TrendsTab.tsx`
- Modify: `stats/src/hooks/useTrends.ts`

**Context:** Add the 6 new trend charts. The hook needs to fetch additional data from the new endpoints.

**Step 1: Extend useTrends hook**

Add fetches for:
- `getEpisodesPerDay(limit)`
- `getNewAnimePerDay(limit)`
- `getWatchTimePerAnime(limit)`

Return these alongside existing data.

**Step 2: Add new TrendChart instances to TrendsTab**

After the existing 9 charts, add:
- Episodes per Day (bar chart)
- Anime Completion Progress (line chart — cumulative)
- New Anime Started (bar chart)
- Watch Time per Anime (stacked bar — needs a new `StackedTrendChart` component or extend `TrendChart`)
- Streak History (visual timeline)
- Cards per Episode (line chart — derive from cards/episodes per day)

For the stacked bar chart, extend `TrendChart` to accept a `stacked` prop with multiple data series, or create a `StackedTrendChart` wrapper.

**Step 3: Organize charts into sections**

Group visually with section headers: "Activity", "Anime", "Efficiency".

**Step 4: Commit**

```bash
git commit -m "feat(stats): add 6 new trend charts for episodes, anime, efficiency"
```

---

## Task 16: Redesign Vocabulary Tab

**Files:**
- Modify: `stats/src/components/vocabulary/VocabularyTab.tsx`
- Modify: `stats/src/components/vocabulary/WordList.tsx`
- Modify: `stats/src/components/vocabulary/KanjiBreakdown.tsx`
- Modify: `stats/src/components/vocabulary/VocabularyOccurrencesDrawer.tsx`
- Create: `stats/src/components/vocabulary/WordDetailPanel.tsx`
- Create: `stats/src/components/vocabulary/KanjiDetailPanel.tsx`

**Context:** The existing drawer shows occurrence lines. Replace/extend it with a full detail panel showing POS, anime appearances, example lines, similar words.

**Step 1: Update VocabularyTab**

- Hero stats: update to use POS-filtered counts (exclude particles)
- Add POS filter toggle (checkbox to show/hide particles, single-char tokens)
- Add search input for word/reading search
- Keep top words chart and new words timeline

**Step 2: Update WordList**

- Show POS tag badge next to each word
- Make each word clickable → opens `WordDetailPanel`
- Support POS filtering from parent

**Step 3: Create WordDetailPanel**

Slide-out panel (reuse `VocabularyOccurrencesDrawer` pattern):
- Header: headword, reading, POS (pos1/pos2/pos3)
- Stats: frequency, first/last seen
- Anime appearances: list of anime with per-anime frequency (from `getWordDetail`)
- Example lines: paginated subtitle lines (from existing `getWordOccurrences`)
- Similar words: clickable list (from `getWordDetail`)

Uses `useWordDetail(wordId)` hook.

**Step 4: Update KanjiBreakdown**

Same pattern: clickable kanji → `KanjiDetailPanel` with frequency, anime appearances, example lines, words using this kanji.

**Step 5: Commit**

```bash
git commit -m "feat(stats): redesign vocabulary tab with POS filter and detail panels"
```

---

## Task 17: Enhance Sessions Tab

**Files:**
- Modify: `stats/src/components/sessions/SessionsTab.tsx`
- Modify: `stats/src/components/sessions/SessionRow.tsx`
- Modify: `stats/src/components/sessions/SessionDetail.tsx`

**Context:** The existing `SessionsTab` is functional but hidden. Enable it and add anime/episode context.

**Step 1: Update SessionRow**

- Add cover art thumbnail (from anime cover endpoint)
- Show anime title + episode number instead of just canonical title
- Keep: duration, cards, lines, lookup rate

**Step 2: Update SessionsTab**

- Add filter by anime title
- Add date range filter
- Group sessions by day (Today / Yesterday / date)

**Step 3: Verify SessionDetail still works**

The inline expandable detail with timeline chart and events should work as-is.

**Step 4: Commit**

```bash
git commit -m "feat(stats): enhance sessions tab with anime context and filters"
```

---

## Task 18: Wire Up Cross-Tab Navigation

**Files:**
- Modify: `stats/src/App.tsx`
- Modify various components

**Context:** Enable navigation between tabs when clicking related items:
- Clicking a word in the Anime detail "words from this anime" should navigate to Vocabulary tab and open that word's detail
- Clicking an anime in the word detail "anime appearances" should navigate to Anime tab and open that anime's detail

**Step 1: Lift navigation state to App level**

Add state for: `selectedAnimeId`, `selectedWordId`. Pass navigation callbacks down to tabs.

**Step 2: Wire up cross-references**

- AnimeWordList: onClick navigates to vocabulary tab + opens word detail
- WordDetailPanel anime appearances: onClick navigates to anime tab + opens anime detail

**Step 3: Commit**

```bash
git commit -m "feat(stats): add cross-tab navigation between anime and vocabulary"
```

---

## Task 19: Final Integration Testing and Polish

**Files:**
- All modified files
- Modify: `stats/src/lib/dashboard-data.test.ts`

**Step 1: Run all backend tests**

Run: `node --test src/core/services/__tests__/stats-server.test.ts`

**Step 2: Run all frontend tests**

Run: `npx vitest run`

**Step 3: Build the stats frontend**

Run: `cd stats && npm run build`

**Step 4: Visual testing**

Start the app and verify each tab renders correctly with real data.

**Step 5: Final commit**

```bash
git commit -m "feat(stats): complete stats dashboard redesign"
```
