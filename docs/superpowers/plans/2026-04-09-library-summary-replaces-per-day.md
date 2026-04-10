# Library Summary Replaces Per-Day Trends — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the six noisy "Library — Per Day" stacked-area charts on the stats Trends tab with a single "Library — Summary" section containing a top-10 watch-time leaderboard and a sortable per-title table, both scoped to the existing date range.

**Architecture:** Backend adds a `librarySummary: LibrarySummaryRow[]` field to the existing `/api/stats/trends/dashboard` response (aggregated from `imm_daily_rollups` + `imm_sessions` joined to `imm_videos`/`imm_anime`) and drops the now-unused `animePerDay` field. Frontend adds a new `LibrarySummarySection` React component (Recharts horizontal bar + sortable HTML table), replaces the per-day section in `TrendsTab.tsx`, and updates all test fixtures.

**Tech Stack:** TypeScript, Bun test runner (backend), Node test runner (frontend), Recharts, React, Tailwind, SQLite (better-sqlite3 via `./sqlite` wrapper).

**Spec:** `docs/superpowers/specs/2026-04-09-library-summary-replaces-per-day-design.md`

---

## File Structure

**Backend (`src/core/services/immersion-tracker/`):**
- `query-trends.ts` — add `LibrarySummaryRow` type, `buildLibrarySummary` helper, wire into `getTrendsDashboard`, drop `animePerDay` from `TrendsDashboardQueryResult`, delete now-unused `buildPerAnimeFromSessions` and `buildLookupsPerHundredPerAnime`.
- `__tests__/query.test.ts` — update existing `getTrendsDashboard` test (drop `animePerDay` assertion, add `librarySummary` assertion); add new tests for summary-specific behavior (empty window, multi-title, null lookupsPerHundred).

**Backend test fixtures:**
- `src/core/services/__tests__/stats-server.test.ts` — update `TRENDS_DASHBOARD` fixture (remove `animePerDay`, add `librarySummary`), fix `assert.deepEqual` that references `body.animePerDay.watchTime`.

**Frontend (`stats/src/`):**
- `types/stats.ts` — add `LibrarySummaryRow` interface, add `librarySummary` field to `TrendsDashboardData`, remove `animePerDay` field.
- `lib/api-client.test.ts` — update the two inline fetch-mock fixtures (remove `animePerDay`, add `librarySummary`).
- `components/trends/LibrarySummarySection.tsx` — **new** file. Owns the header content: leaderboard Recharts chart + sortable HTML table. Takes `{ rows, hiddenTitles }` as props.
- `components/trends/TrendsTab.tsx` — delete the "Library — Per Day" block (lines 224-254 and the filtered data locals 137-146), add `LibrarySummarySection` import + usage, update `buildAnimeVisibilityOptions` call to use `librarySummary` titles instead of the six dropped `animePerDay.*` arrays.
- `components/trends/anime-visibility.ts` — unchanged. The existing helpers operate on `PerAnimeDataPoint[]`; we'll adapt by passing a derived `PerAnimeDataPoint[]` built from `librarySummary` (or add an overload — see Task 7 for the final decision).

**Changelog:**
- `changes/stats-library-summary.md` — **new** changelog fragment.

---

## Task 1: Backend — Add `LibrarySummaryRow` type and empty stub field

**Files:**
- Modify: `src/core/services/immersion-tracker/query-trends.ts`

- [ ] **Step 1: Add the row type and add `librarySummary: []` to the returned object**

Edit `src/core/services/immersion-tracker/query-trends.ts`. After the existing `TrendPerAnimePoint` interface (around line 24), add:

```ts
export interface LibrarySummaryRow {
  title: string;
  watchTimeMin: number;
  videos: number;
  sessions: number;
  cards: number;
  words: number;
  lookups: number;
  lookupsPerHundred: number | null;
  firstWatched: number;
  lastWatched: number;
}
```

In the same file, add a new field to `TrendsDashboardQueryResult` (around line 45-82), alongside `animePerDay`:

```ts
librarySummary: LibrarySummaryRow[];
```

In `getTrendsDashboard` (around line 622), add `librarySummary: []` to the returned object literal (inside the final `return { ... }`). Keep everything else as-is for now.

- [ ] **Step 2: Run typecheck**

Run: `bun run typecheck`
Expected: PASS (empty array satisfies the new field; no downstream consumer yet).

- [ ] **Step 3: Commit**

```bash
git add src/core/services/immersion-tracker/query-trends.ts
git commit -m "feat(stats): scaffold LibrarySummaryRow type and empty field"
```

---

## Task 2: Backend — TDD the `buildLibrarySummary` helper

**Files:**
- Modify: `src/core/services/immersion-tracker/query-trends.ts`
- Modify: `src/core/services/immersion-tracker/__tests__/query.test.ts`

- [ ] **Step 1: Write a failing unit test for the happy path**

Open `src/core/services/immersion-tracker/__tests__/query.test.ts` and add a new test at the end of the file (before the last closing brace, or after the last `test(...)` block — verify by reading the end of the file). Use the same imports/helpers the existing `getTrendsDashboard` tests use (`makeDbPath`, `ensureSchema`, `getOrCreateVideoRecord`, `getOrCreateAnimeRecord`, `linkVideoToAnimeRecord`, `startSessionRecord`, `createTrackerPreparedStatements`, `Database`, `cleanupDbPath`, `getTrendsDashboard`, `SOURCE_TYPE_LOCAL`).

```ts
test('getTrendsDashboard builds librarySummary with per-title aggregates', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/library-summary-test.mkv', {
      canonicalTitle: 'Library Summary Test',
      sourcePath: '/tmp/library-summary-test.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Summary Anime',
      canonicalTitle: 'Summary Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: 'library-summary-test.mkv',
      parsedTitle: 'Summary Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'test',
      parserConfidence: 1,
      parseMetadataJson: null,
    });

    const dayOneStart = 1_700_000_000_000;
    const dayTwoStart = dayOneStart + 86_400_000;

    const sessionOne = startSessionRecord(db, videoId, dayOneStart);
    const sessionTwo = startSessionRecord(db, videoId, dayTwoStart);

    for (const [sessionId, startedAtMs, activeMs, cards, tokens, lookups] of [
      [sessionOne.sessionId, dayOneStart, 30 * 60_000, 2, 120, 8],
      [sessionTwo.sessionId, dayTwoStart, 45 * 60_000, 3, 140, 10],
    ] as const) {
      stmts.telemetryInsertStmt.run(
        sessionId,
        `${startedAtMs + 60_000}`,
        activeMs,
        activeMs,
        10,
        tokens,
        cards,
        0,
        0,
        lookups,
        0,
        0,
        0,
        0,
        `${startedAtMs + 60_000}`,
        `${startedAtMs + 60_000}`,
      );

      db.prepare(
        `
          UPDATE imm_sessions
          SET ended_at_ms = ?, total_watched_ms = ?, active_watched_ms = ?,
              lines_seen = ?, tokens_seen = ?, cards_mined = ?, yomitan_lookup_count = ?
          WHERE session_id = ?
        `,
      ).run(
        `${startedAtMs + activeMs}`,
        activeMs,
        activeMs,
        10,
        tokens,
        cards,
        lookups,
        sessionId,
      );
    }

    for (const [day, active, tokens, cards] of [
      [Math.floor(dayOneStart / 86_400_000), 30, 120, 2],
      [Math.floor(dayTwoStart / 86_400_000), 45, 140, 3],
    ] as const) {
      db.prepare(
        `
          INSERT INTO imm_daily_rollups (
            rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
            total_tokens_seen, total_cards
          ) VALUES (?, ?, ?, ?, ?, ?, ?)
        `,
      ).run(day, videoId, 1, active, 10, tokens, cards);
    }

    const dashboard = getTrendsDashboard(db, 'all', 'day');

    assert.equal(dashboard.librarySummary.length, 1);
    const row = dashboard.librarySummary[0]!;
    assert.equal(row.title, 'Summary Anime');
    assert.equal(row.watchTimeMin, 75);
    assert.equal(row.videos, 1);
    assert.equal(row.sessions, 2);
    assert.equal(row.cards, 5);
    assert.equal(row.words, 260);
    assert.equal(row.lookups, 18);
    assert.equal(row.lookupsPerHundred, +((18 / 260) * 100).toFixed(1));
    assert.equal(row.firstWatched, Math.floor(dayOneStart / 86_400_000));
    assert.equal(row.lastWatched, Math.floor(dayTwoStart / 86_400_000));
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts -t "librarySummary with per-title aggregates"`
Expected: FAIL — `dashboard.librarySummary.length` is `0`, not `1`.

- [ ] **Step 3: Implement the `buildLibrarySummary` helper**

Open `src/core/services/immersion-tracker/query-trends.ts`. Add this helper function near the other builders (e.g., after `buildCumulativePerAnime`, before `getVideoAnimeTitleMap`):

```ts
function buildLibrarySummary(
  rollups: ImmersionSessionRollupRow[],
  sessions: TrendSessionMetricRow[],
  titlesByVideoId: Map<number, string>,
): LibrarySummaryRow[] {
  type Accum = {
    watchTimeMin: number;
    videos: Set<number>;
    cards: number;
    words: number;
    firstWatched: number;
    lastWatched: number;
    sessions: number;
    lookups: number;
  };

  const byTitle = new Map<string, Accum>();

  const ensure = (title: string): Accum => {
    const existing = byTitle.get(title);
    if (existing) return existing;
    const created: Accum = {
      watchTimeMin: 0,
      videos: new Set<number>(),
      cards: 0,
      words: 0,
      firstWatched: Number.POSITIVE_INFINITY,
      lastWatched: Number.NEGATIVE_INFINITY,
      sessions: 0,
      lookups: 0,
    };
    byTitle.set(title, created);
    return created;
  };

  for (const rollup of rollups) {
    if (rollup.videoId === null) continue;
    const title = resolveVideoAnimeTitle(rollup.videoId, titlesByVideoId);
    const acc = ensure(title);
    acc.watchTimeMin += rollup.totalActiveMin;
    acc.cards += rollup.totalCards;
    acc.words += rollup.totalTokensSeen;
    acc.videos.add(rollup.videoId);
    if (rollup.rollupDayOrMonth < acc.firstWatched) {
      acc.firstWatched = rollup.rollupDayOrMonth;
    }
    if (rollup.rollupDayOrMonth > acc.lastWatched) {
      acc.lastWatched = rollup.rollupDayOrMonth;
    }
  }

  for (const session of sessions) {
    const title = resolveTrendAnimeTitle(session);
    if (!byTitle.has(title)) continue;
    const acc = byTitle.get(title)!;
    acc.sessions += 1;
    acc.lookups += session.yomitanLookupCount;
  }

  const rows: LibrarySummaryRow[] = [];
  for (const [title, acc] of byTitle) {
    if (!Number.isFinite(acc.firstWatched) || !Number.isFinite(acc.lastWatched)) {
      continue;
    }
    rows.push({
      title,
      watchTimeMin: Math.round(acc.watchTimeMin),
      videos: acc.videos.size,
      sessions: acc.sessions,
      cards: acc.cards,
      words: acc.words,
      lookups: acc.lookups,
      lookupsPerHundred:
        acc.words > 0 ? +((acc.lookups / acc.words) * 100).toFixed(1) : null,
      firstWatched: acc.firstWatched,
      lastWatched: acc.lastWatched,
    });
  }

  rows.sort((a, b) => b.watchTimeMin - a.watchTimeMin || a.title.localeCompare(b.title));
  return rows;
}
```

- [ ] **Step 4: Wire `buildLibrarySummary` into `getTrendsDashboard`**

Still in `query-trends.ts`, inside `getTrendsDashboard`, replace the stub `librarySummary: []` line with:

```ts
librarySummary: buildLibrarySummary(dailyRollups, sessions, titlesByVideoId),
```

Place it at the same spot in the return object (keep existing fields otherwise unchanged).

- [ ] **Step 5: Run the test to verify it passes**

Run: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts -t "librarySummary with per-title aggregates"`
Expected: PASS.

- [ ] **Step 6: Run the full query test file to ensure no regressions**

Run: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`
Expected: PASS for all tests.

- [ ] **Step 7: Commit**

```bash
git add src/core/services/immersion-tracker/query-trends.ts \
        src/core/services/immersion-tracker/__tests__/query.test.ts
git commit -m "feat(stats): build per-title librarySummary from daily rollups and sessions"
```

---

## Task 3: Backend — Add null-lookupsPerHundred and empty-window tests

**Files:**
- Modify: `src/core/services/immersion-tracker/__tests__/query.test.ts`

- [ ] **Step 1: Write a failing test for `lookupsPerHundred: null` when words == 0**

Append to `src/core/services/immersion-tracker/__tests__/query.test.ts`:

```ts
test('getTrendsDashboard librarySummary returns null lookupsPerHundred when words is zero', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/lib-summary-null.mkv', {
      canonicalTitle: 'Null Lookups Title',
      sourcePath: '/tmp/lib-summary-null.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Null Lookups Anime',
      canonicalTitle: 'Null Lookups Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: 'lib-summary-null.mkv',
      parsedTitle: 'Null Lookups Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'test',
      parserConfidence: 1,
      parseMetadataJson: null,
    });

    const startMs = 1_700_000_000_000;
    const session = startSessionRecord(db, videoId, startMs);
    stmts.telemetryInsertStmt.run(
      session.sessionId,
      `${startMs + 60_000}`,
      20 * 60_000,
      20 * 60_000,
      5,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      `${startMs + 60_000}`,
      `${startMs + 60_000}`,
    );
    db.prepare(
      `
        UPDATE imm_sessions
        SET ended_at_ms = ?, total_watched_ms = ?, active_watched_ms = ?,
            lines_seen = ?, tokens_seen = ?, cards_mined = ?, yomitan_lookup_count = ?
        WHERE session_id = ?
      `,
    ).run(
      `${startMs + 20 * 60_000}`,
      20 * 60_000,
      20 * 60_000,
      5,
      0,
      0,
      0,
      session.sessionId,
    );

    db.prepare(
      `
        INSERT INTO imm_daily_rollups (
          rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
          total_tokens_seen, total_cards
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(Math.floor(startMs / 86_400_000), videoId, 1, 20, 5, 0, 0);

    const dashboard = getTrendsDashboard(db, 'all', 'day');
    assert.equal(dashboard.librarySummary.length, 1);
    assert.equal(dashboard.librarySummary[0]!.lookupsPerHundred, null);
    assert.equal(dashboard.librarySummary[0]!.words, 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getTrendsDashboard librarySummary is empty when no rollups exist', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const dashboard = getTrendsDashboard(db, 'all', 'day');
    assert.deepEqual(dashboard.librarySummary, []);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});
```

- [ ] **Step 2: Run the new tests**

Run: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts -t "librarySummary"`
Expected: PASS for all three librarySummary tests (the helper implemented in Task 2 already handles these cases).

- [ ] **Step 3: Commit**

```bash
git add src/core/services/immersion-tracker/__tests__/query.test.ts
git commit -m "test(stats): cover librarySummary null-lookups and empty-window cases"
```

---

## Task 4: Backend — Drop `animePerDay` from the response type and clean up dead helpers

**Files:**
- Modify: `src/core/services/immersion-tracker/query-trends.ts`
- Modify: `src/core/services/immersion-tracker/__tests__/query.test.ts`
- Modify: `src/core/services/__tests__/stats-server.test.ts`

- [ ] **Step 1: Remove `animePerDay` from `TrendsDashboardQueryResult`**

In `src/core/services/immersion-tracker/query-trends.ts`, delete the `animePerDay` block from the interface (lines ~64-71):

```ts
// Delete this block:
animePerDay: {
  episodes: TrendPerAnimePoint[];
  watchTime: TrendPerAnimePoint[];
  cards: TrendPerAnimePoint[];
  words: TrendPerAnimePoint[];
  lookups: TrendPerAnimePoint[];
  lookupsPerHundred: TrendPerAnimePoint[];
};
```

- [ ] **Step 2: Scope the intermediate `animePerDay` to a local variable and drop it from the return**

In `getTrendsDashboard` (around lines 649-668 and 694-699), keep the internal `animePerDay` construction (it's still used by `animeCumulative`) but do NOT include it in the returned object. Also drop the now-unused `lookups` and `lookupsPerHundred` fields from the internal `animePerDay` object. Replace the block starting with `const animePerDay = {` through the return statement:

```ts
  const animePerDay = {
    episodes: buildEpisodesPerAnimeFromDailyRollups(dailyRollups, titlesByVideoId),
    watchTime: buildPerAnimeFromDailyRollups(
      dailyRollups,
      titlesByVideoId,
      (rollup) => rollup.totalActiveMin,
    ),
    cards: buildPerAnimeFromDailyRollups(
      dailyRollups,
      titlesByVideoId,
      (rollup) => rollup.totalCards,
    ),
    words: buildPerAnimeFromDailyRollups(
      dailyRollups,
      titlesByVideoId,
      (rollup) => rollup.totalTokensSeen,
    ),
  };

  return {
    activity,
    progress: {
      watchTime: accumulatePoints(activity.watchTime),
      sessions: accumulatePoints(activity.sessions),
      words: accumulatePoints(activity.words),
      newWords: accumulatePoints(
        useMonthlyBuckets ? buildNewWordsPerMonth(db, cutoffMs) : buildNewWordsPerDay(db, cutoffMs),
      ),
      cards: accumulatePoints(activity.cards),
      episodes: accumulatePoints(
        useMonthlyBuckets
          ? buildEpisodesPerMonthFromRollups(monthlyRollups)
          : buildEpisodesPerDayFromDailyRollups(dailyRollups),
      ),
      lookups: accumulatePoints(
        useMonthlyBuckets
          ? buildSessionSeriesByMonth(sessions, (session) => session.yomitanLookupCount)
          : buildSessionSeriesByDay(sessions, (session) => session.yomitanLookupCount),
      ),
    },
    ratios: {
      lookupsPerHundred: buildLookupsPerHundredWords(sessions, groupBy),
    },
    librarySummary: buildLibrarySummary(dailyRollups, sessions, titlesByVideoId),
    animeCumulative: {
      watchTime: buildCumulativePerAnime(animePerDay.watchTime),
      episodes: buildCumulativePerAnime(animePerDay.episodes),
      cards: buildCumulativePerAnime(animePerDay.cards),
      words: buildCumulativePerAnime(animePerDay.words),
    },
    patterns: {
      watchTimeByDayOfWeek: buildWatchTimeByDayOfWeek(sessions),
      watchTimeByHour: buildWatchTimeByHour(sessions),
    },
  };
```

- [ ] **Step 3: Delete now-unused helpers**

In the same file, delete the functions `buildPerAnimeFromSessions` (around lines 304-325) and `buildLookupsPerHundredPerAnime` (around lines 327-357). Nothing else references them after Step 2.

- [ ] **Step 4: Update the existing `getTrendsDashboard` test assertion that references `animePerDay`**

In `src/core/services/immersion-tracker/__tests__/query.test.ts`, find the test `getTrendsDashboard returns chart-ready aggregated series`. Replace the line:

```ts
assert.equal(dashboard.animePerDay.watchTime[0]?.animeTitle, 'Trend Dashboard Anime');
```

with:

```ts
assert.equal(dashboard.librarySummary[0]?.title, 'Trend Dashboard Anime');
```

- [ ] **Step 5: Update the stats-server test fixture**

In `src/core/services/__tests__/stats-server.test.ts`, find `TRENDS_DASHBOARD` (around line 150). Remove the entire `animePerDay: { ... }` block (lines ~169-176). Add a `librarySummary` field inside the fixture (anywhere appropriate — before `animeCumulative` is fine):

```ts
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
```

Then find the assertion around line 601:

```ts
assert.deepEqual(body.animePerDay.watchTime, TRENDS_DASHBOARD.animePerDay.watchTime);
```

Replace it with:

```ts
assert.deepEqual(body.librarySummary, TRENDS_DASHBOARD.librarySummary);
```

- [ ] **Step 6: Run typecheck + backend tests**

Run: `bun run typecheck`
Expected: PASS.

Run: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`
Expected: PASS.

Run: `bun test src/core/services/__tests__/stats-server.test.ts`
Expected: PASS.

- [ ] **Step 7: Commit**

```bash
git add src/core/services/immersion-tracker/query-trends.ts \
        src/core/services/immersion-tracker/__tests__/query.test.ts \
        src/core/services/__tests__/stats-server.test.ts
git commit -m "refactor(stats): drop animePerDay from trends response in favor of librarySummary"
```

---

## Task 5: Frontend — Update types and api-client test fixtures

**Files:**
- Modify: `stats/src/types/stats.ts`
- Modify: `stats/src/lib/api-client.test.ts`

- [ ] **Step 1: Add `LibrarySummaryRow` and update `TrendsDashboardData` in `stats/src/types/stats.ts`**

Add above `TrendsDashboardData` (around line 291):

```ts
export interface LibrarySummaryRow {
  title: string;
  watchTimeMin: number;
  videos: number;
  sessions: number;
  cards: number;
  words: number;
  lookups: number;
  lookupsPerHundred: number | null;
  firstWatched: number;
  lastWatched: number;
}
```

Inside `TrendsDashboardData`, delete the `animePerDay` block (lines ~310-317) and add `librarySummary: LibrarySummaryRow[];` (place it before `animeCumulative`):

```ts
export interface TrendsDashboardData {
  activity: {
    watchTime: TrendChartPoint[];
    cards: TrendChartPoint[];
    words: TrendChartPoint[];
    sessions: TrendChartPoint[];
  };
  progress: {
    watchTime: TrendChartPoint[];
    sessions: TrendChartPoint[];
    words: TrendChartPoint[];
    newWords: TrendChartPoint[];
    cards: TrendChartPoint[];
    episodes: TrendChartPoint[];
    lookups: TrendChartPoint[];
  };
  ratios: {
    lookupsPerHundred: TrendChartPoint[];
  };
  librarySummary: LibrarySummaryRow[];
  animeCumulative: {
    watchTime: TrendPerAnimePoint[];
    episodes: TrendPerAnimePoint[];
    cards: TrendPerAnimePoint[];
    words: TrendPerAnimePoint[];
  };
  patterns: {
    watchTimeByDayOfWeek: TrendChartPoint[];
    watchTimeByHour: TrendChartPoint[];
  };
}
```

- [ ] **Step 2: Update the two inline fixtures in `stats/src/lib/api-client.test.ts`**

Find both inline `JSON.stringify({ ... })` fetch-mock bodies (around lines 75-107 and 123-150). In **both** blocks, delete the `animePerDay: { ... }` object and replace with:

```ts
librarySummary: [],
```

(Insert before `animeCumulative`.)

- [ ] **Step 3: Run frontend typecheck and tests**

Run: `cd stats && bun run typecheck`
Expected: FAIL — `TrendsTab.tsx` still references `data.animePerDay`. That's expected; we fix it in Task 8. Continue.

Run: `cd stats && bun test src/lib/api-client.test.ts`
Expected: PASS (the test only asserts URL construction, not response shape).

- [ ] **Step 4: Commit**

```bash
git add stats/src/types/stats.ts stats/src/lib/api-client.test.ts
git commit -m "refactor(stats): replace animePerDay type with librarySummary"
```

---

## Task 6: Frontend — Create `LibrarySummarySection` skeleton with empty state

**Files:**
- Create: `stats/src/components/trends/LibrarySummarySection.tsx`

- [ ] **Step 1: Create the file with the empty state and props plumbing**

Create `stats/src/components/trends/LibrarySummarySection.tsx`:

```tsx
import type { LibrarySummaryRow } from '../../types/stats';

interface LibrarySummarySectionProps {
  rows: LibrarySummaryRow[];
  hiddenTitles: ReadonlySet<string>;
}

export function LibrarySummarySection({ rows, hiddenTitles }: LibrarySummarySectionProps) {
  const visibleRows = rows.filter((row) => !hiddenTitles.has(row.title));

  if (visibleRows.length === 0) {
    return (
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <div className="text-xs text-ctp-overlay2">No library activity in the selected window.</div>
      </div>
    );
  }

  return (
    <>
      {/* Leaderboard + table cards added in Tasks 7 and 8 */}
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <div className="text-xs text-ctp-overlay2">
          Library summary: {visibleRows.length} titles
        </div>
      </div>
    </>
  );
}
```

- [ ] **Step 2: Run typecheck (it will still fail in `TrendsTab.tsx`, but the new file should typecheck cleanly)**

Run: `cd stats && bun run typecheck 2>&1 | grep -E 'LibrarySummarySection\.tsx'`
Expected: no output (new file has no type errors). `TrendsTab.tsx` still errors — ignore until Task 8.

- [ ] **Step 3: Commit**

```bash
git add stats/src/components/trends/LibrarySummarySection.tsx
git commit -m "feat(stats): scaffold LibrarySummarySection with empty state"
```

---

## Task 7: Frontend — Add the leaderboard bar chart to `LibrarySummarySection`

**Files:**
- Modify: `stats/src/components/trends/LibrarySummarySection.tsx`

- [ ] **Step 1: Replace the skeleton body with the leaderboard chart**

Replace the entire contents of `stats/src/components/trends/LibrarySummarySection.tsx` with:

```tsx
import {
  Bar,
  BarChart,
  CartesianGrid,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import type { LibrarySummaryRow } from '../../types/stats';
import { CHART_DEFAULTS, CHART_THEME, TOOLTIP_CONTENT_STYLE } from '../../lib/chart-theme';

interface LibrarySummarySectionProps {
  rows: LibrarySummaryRow[];
  hiddenTitles: ReadonlySet<string>;
}

const LEADERBOARD_LIMIT = 10;
const LEADERBOARD_HEIGHT = 260;
const LEADERBOARD_BAR_COLOR = '#8aadf4';

function truncateTitle(title: string, maxChars: number): string {
  if (title.length <= maxChars) return title;
  return `${title.slice(0, maxChars - 1)}…`;
}

export function LibrarySummarySection({ rows, hiddenTitles }: LibrarySummarySectionProps) {
  const visibleRows = rows.filter((row) => !hiddenTitles.has(row.title));

  if (visibleRows.length === 0) {
    return (
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <div className="text-xs text-ctp-overlay2">No library activity in the selected window.</div>
      </div>
    );
  }

  const leaderboard = [...visibleRows]
    .sort((a, b) => b.watchTimeMin - a.watchTimeMin)
    .slice(0, LEADERBOARD_LIMIT)
    .map((row) => ({
      title: row.title,
      displayTitle: truncateTitle(row.title, 24),
      watchTimeMin: row.watchTimeMin,
    }));

  return (
    <>
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <h3 className="text-xs font-semibold text-ctp-text mb-2">
          Top Titles by Watch Time (min)
        </h3>
        <ResponsiveContainer width="100%" height={LEADERBOARD_HEIGHT}>
          <BarChart
            data={leaderboard}
            layout="vertical"
            margin={{ top: 8, right: 16, bottom: 8, left: 8 }}
          >
            <CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />
            <XAxis
              type="number"
              tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
              axisLine={{ stroke: CHART_THEME.axisLine }}
              tickLine={false}
            />
            <YAxis
              type="category"
              dataKey="displayTitle"
              width={160}
              tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
              axisLine={{ stroke: CHART_THEME.axisLine }}
              tickLine={false}
              interval={0}
            />
            <Tooltip
              contentStyle={TOOLTIP_CONTENT_STYLE}
              formatter={(value: number) => [`${value} min`, 'Watch Time']}
              labelFormatter={(_label, payload) => {
                const datum = payload?.[0]?.payload as { title?: string } | undefined;
                return datum?.title ?? '';
              }}
            />
            <Bar dataKey="watchTimeMin" fill={LEADERBOARD_BAR_COLOR} />
          </BarChart>
        </ResponsiveContainer>
      </div>
      {/* Table card added in Task 8 */}
    </>
  );
}
```

- [ ] **Step 2: Typecheck (component in isolation)**

Run: `cd stats && bun run typecheck 2>&1 | grep -E 'LibrarySummarySection\.tsx'`
Expected: no output — new component typechecks. `TrendsTab.tsx` errors remain (fixed in Task 9).

- [ ] **Step 3: Commit**

```bash
git add stats/src/components/trends/LibrarySummarySection.tsx
git commit -m "feat(stats): add top-titles leaderboard chart to LibrarySummarySection"
```

---

## Task 8: Frontend — Add the sortable table to `LibrarySummarySection`

**Files:**
- Modify: `stats/src/components/trends/LibrarySummarySection.tsx`

- [ ] **Step 1: Add sort state, column definitions, and the table markup**

Replace the entire file with the version below. The change vs. Task 7: imports `useState`, `useMemo`, and `formatDuration` + `epochDayToDate`; adds `SortColumn`, `SortDirection`, `COLUMNS`; adds a `<table>` card after the leaderboard card.

```tsx
import { useMemo, useState } from 'react';
import {
  Bar,
  BarChart,
  CartesianGrid,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import type { LibrarySummaryRow } from '../../types/stats';
import { CHART_DEFAULTS, CHART_THEME, TOOLTIP_CONTENT_STYLE } from '../../lib/chart-theme';
import { epochDayToDate, formatDuration, formatNumber } from '../../lib/formatters';

interface LibrarySummarySectionProps {
  rows: LibrarySummaryRow[];
  hiddenTitles: ReadonlySet<string>;
}

const LEADERBOARD_LIMIT = 10;
const LEADERBOARD_HEIGHT = 260;
const LEADERBOARD_BAR_COLOR = '#8aadf4';
const TABLE_MAX_HEIGHT = 480;

type SortColumn =
  | 'title'
  | 'watchTimeMin'
  | 'videos'
  | 'sessions'
  | 'cards'
  | 'words'
  | 'lookups'
  | 'lookupsPerHundred'
  | 'firstWatched';

type SortDirection = 'asc' | 'desc';

interface ColumnDef {
  id: SortColumn;
  label: string;
  align: 'left' | 'right';
}

const COLUMNS: ColumnDef[] = [
  { id: 'title', label: 'Title', align: 'left' },
  { id: 'watchTimeMin', label: 'Watch Time', align: 'right' },
  { id: 'videos', label: 'Videos', align: 'right' },
  { id: 'sessions', label: 'Sessions', align: 'right' },
  { id: 'cards', label: 'Cards', align: 'right' },
  { id: 'words', label: 'Words', align: 'right' },
  { id: 'lookups', label: 'Lookups', align: 'right' },
  { id: 'lookupsPerHundred', label: 'Lookups/100w', align: 'right' },
  { id: 'firstWatched', label: 'Date Range', align: 'right' },
];

function truncateTitle(title: string, maxChars: number): string {
  if (title.length <= maxChars) return title;
  return `${title.slice(0, maxChars - 1)}…`;
}

function formatDateRange(firstEpochDay: number, lastEpochDay: number): string {
  const fmt = (epochDay: number) =>
    epochDayToDate(epochDay).toLocaleDateString(undefined, {
      month: 'short',
      day: 'numeric',
    });
  if (firstEpochDay === lastEpochDay) return fmt(firstEpochDay);
  return `${fmt(firstEpochDay)} → ${fmt(lastEpochDay)}`;
}

function formatWatchTime(min: number): string {
  return formatDuration(min * 60_000);
}

function compareRows(
  a: LibrarySummaryRow,
  b: LibrarySummaryRow,
  column: SortColumn,
  direction: SortDirection,
): number {
  const sign = direction === 'asc' ? 1 : -1;

  if (column === 'title') {
    return a.title.localeCompare(b.title) * sign;
  }

  if (column === 'firstWatched') {
    return (a.firstWatched - b.firstWatched) * sign;
  }

  if (column === 'lookupsPerHundred') {
    // Null sorts as lowest in both directions (treated as "no data").
    const aVal = a.lookupsPerHundred;
    const bVal = b.lookupsPerHundred;
    if (aVal === null && bVal === null) return 0;
    if (aVal === null) return 1;
    if (bVal === null) return -1;
    return (aVal - bVal) * sign;
  }

  const aVal = a[column] as number;
  const bVal = b[column] as number;
  return (aVal - bVal) * sign;
}

export function LibrarySummarySection({ rows, hiddenTitles }: LibrarySummarySectionProps) {
  const [sortColumn, setSortColumn] = useState<SortColumn>('watchTimeMin');
  const [sortDirection, setSortDirection] = useState<SortDirection>('desc');

  const visibleRows = useMemo(
    () => rows.filter((row) => !hiddenTitles.has(row.title)),
    [rows, hiddenTitles],
  );

  const sortedRows = useMemo(
    () => [...visibleRows].sort((a, b) => compareRows(a, b, sortColumn, sortDirection)),
    [visibleRows, sortColumn, sortDirection],
  );

  const leaderboard = useMemo(
    () =>
      [...visibleRows]
        .sort((a, b) => b.watchTimeMin - a.watchTimeMin)
        .slice(0, LEADERBOARD_LIMIT)
        .map((row) => ({
          title: row.title,
          displayTitle: truncateTitle(row.title, 24),
          watchTimeMin: row.watchTimeMin,
        })),
    [visibleRows],
  );

  if (visibleRows.length === 0) {
    return (
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <div className="text-xs text-ctp-overlay2">
          No library activity in the selected window.
        </div>
      </div>
    );
  }

  const handleHeaderClick = (column: SortColumn) => {
    if (column === sortColumn) {
      setSortDirection((prev) => (prev === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortColumn(column);
      setSortDirection(column === 'title' ? 'asc' : 'desc');
    }
  };

  return (
    <>
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <h3 className="text-xs font-semibold text-ctp-text mb-2">
          Top Titles by Watch Time (min)
        </h3>
        <ResponsiveContainer width="100%" height={LEADERBOARD_HEIGHT}>
          <BarChart
            data={leaderboard}
            layout="vertical"
            margin={{ top: 8, right: 16, bottom: 8, left: 8 }}
          >
            <CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />
            <XAxis
              type="number"
              tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
              axisLine={{ stroke: CHART_THEME.axisLine }}
              tickLine={false}
            />
            <YAxis
              type="category"
              dataKey="displayTitle"
              width={160}
              tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
              axisLine={{ stroke: CHART_THEME.axisLine }}
              tickLine={false}
              interval={0}
            />
            <Tooltip
              contentStyle={TOOLTIP_CONTENT_STYLE}
              formatter={(value: number) => [`${value} min`, 'Watch Time']}
              labelFormatter={(_label, payload) => {
                const datum = payload?.[0]?.payload as { title?: string } | undefined;
                return datum?.title ?? '';
              }}
            />
            <Bar dataKey="watchTimeMin" fill={LEADERBOARD_BAR_COLOR} />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <h3 className="text-xs font-semibold text-ctp-text mb-2">Per-Title Summary</h3>
        <div
          className="overflow-auto"
          style={{ maxHeight: TABLE_MAX_HEIGHT }}
        >
          <table className="w-full text-xs">
            <thead className="sticky top-0 bg-ctp-surface0">
              <tr className="border-b border-ctp-surface1 text-ctp-subtext0">
                {COLUMNS.map((column) => {
                  const isActive = column.id === sortColumn;
                  const indicator = isActive ? (sortDirection === 'asc' ? ' ▲' : ' ▼') : '';
                  return (
                    <th
                      key={column.id}
                      scope="col"
                      className={`px-2 py-2 font-medium select-none cursor-pointer hover:text-ctp-text ${
                        column.align === 'right' ? 'text-right' : 'text-left'
                      } ${isActive ? 'text-ctp-text' : ''}`}
                      onClick={() => handleHeaderClick(column.id)}
                    >
                      {column.label}
                      {indicator}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {sortedRows.map((row) => (
                <tr
                  key={row.title}
                  className="border-b border-ctp-surface1 last:border-b-0 hover:bg-ctp-surface1/40"
                >
                  <td
                    className="px-2 py-2 text-left text-ctp-text max-w-[240px] truncate"
                    title={row.title}
                  >
                    {row.title}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatWatchTime(row.watchTimeMin)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.videos)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.sessions)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.cards)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.words)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.lookups)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {row.lookupsPerHundred === null
                      ? '—'
                      : row.lookupsPerHundred.toFixed(1)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-subtext0 tabular-nums">
                    {formatDateRange(row.firstWatched, row.lastWatched)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
}
```

- [ ] **Step 2: Typecheck the new component**

Run: `cd stats && bun run typecheck 2>&1 | grep -E 'LibrarySummarySection\.tsx'`
Expected: no output. `TrendsTab.tsx` errors still remain — next task fixes them.

- [ ] **Step 3: Commit**

```bash
git add stats/src/components/trends/LibrarySummarySection.tsx
git commit -m "feat(stats): add sortable per-title table to LibrarySummarySection"
```

---

## Task 9: Frontend — Wire `LibrarySummarySection` into `TrendsTab` and remove the per-day block

**Files:**
- Modify: `stats/src/components/trends/TrendsTab.tsx`

- [ ] **Step 1: Delete the per-day filtered locals and imports**

In `stats/src/components/trends/TrendsTab.tsx`:

Delete these locals (currently lines ~129-146):

```ts
const filteredEpisodesPerAnime = filterHiddenAnimeData(
  data.animePerDay.episodes,
  activeHiddenAnime,
);
const filteredWatchTimePerAnime = filterHiddenAnimeData(
  data.animePerDay.watchTime,
  activeHiddenAnime,
);
const filteredCardsPerAnime = filterHiddenAnimeData(data.animePerDay.cards, activeHiddenAnime);
const filteredWordsPerAnime = filterHiddenAnimeData(data.animePerDay.words, activeHiddenAnime);
const filteredLookupsPerAnime = filterHiddenAnimeData(
  data.animePerDay.lookups,
  activeHiddenAnime,
);
const filteredLookupsPerHundredPerAnime = filterHiddenAnimeData(
  data.animePerDay.lookupsPerHundred,
  activeHiddenAnime,
);
```

- [ ] **Step 2: Update `buildAnimeVisibilityOptions` to use `librarySummary` titles**

Replace the existing `const animeTitles = buildAnimeVisibilityOptions([...])` block (currently lines 116-126) with:

```ts
const librarySummaryAsPoints = data.librarySummary.map((row) => ({
  epochDay: 0,
  animeTitle: row.title,
  value: row.watchTimeMin,
}));

const animeTitles = buildAnimeVisibilityOptions([
  librarySummaryAsPoints,
  data.animeCumulative.episodes,
  data.animeCumulative.cards,
  data.animeCumulative.words,
  data.animeCumulative.watchTime,
]);
```

This reuses the existing `PerAnimeDataPoint`-shaped helper without modifying it — the `epochDay: 0` is a placeholder the helper never inspects.

- [ ] **Step 3: Import `LibrarySummarySection` at the top of the file**

Add to the imports at the top (near the other `./` imports on line 5):

```ts
import { LibrarySummarySection } from './LibrarySummarySection';
```

- [ ] **Step 4: Replace the "Library — Per Day" JSX block**

Find lines 224-254 (the block starting with `<SectionHeader>Library — Per Day</SectionHeader>`). Replace the entire block through the final `/>` of `Lookups/100w per Title` with:

```tsx
<SectionHeader>Library — Summary</SectionHeader>
<AnimeVisibilityFilter
  animeTitles={animeTitles}
  hiddenAnime={activeHiddenAnime}
  onShowAll={() => setHiddenAnime(new Set())}
  onHideAll={() => setHiddenAnime(new Set(animeTitles))}
  onToggleAnime={(title) =>
    setHiddenAnime((current) => {
      const next = new Set(current);
      if (next.has(title)) {
        next.delete(title);
      } else {
        next.add(title);
      }
      return next;
    })
  }
/>
<LibrarySummarySection rows={data.librarySummary} hiddenTitles={activeHiddenAnime} />
```

(The `AnimeVisibilityFilter` moves from the per-day section into the summary section — same component, same props pattern.)

- [ ] **Step 5: Verify `StackedTrendChart` and `filterHiddenAnimeData` are still imported**

Those imports are still needed by the "Library — Cumulative" section (lines 256-264 — make sure you did NOT delete them). If the linter reports them as unused, they aren't. Do not touch them.

- [ ] **Step 6: Run frontend typecheck**

Run: `cd stats && bun run typecheck`
Expected: PASS (no more `animePerDay` references).

- [ ] **Step 7: Run the full fast test suite**

Run: `bun run test:fast`
Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add stats/src/components/trends/TrendsTab.tsx
git commit -m "feat(stats): replace per-day trends section with library summary"
```

---

## Task 10: Add changelog fragment and run the full handoff gate

**Files:**
- Create: `changes/stats-library-summary.md`

- [ ] **Step 1: Check the existing changelog fragment format**

Run: `ls changes/ && head -20 changes/*.md | head -60`
Inspect a recent fragment to match the exact format (frontmatter, section headings). Base the new fragment on whatever convention you see — do not guess.

- [ ] **Step 2: Write the fragment**

Create `changes/stats-library-summary.md` using the format you just observed. The body should say something like:

> Replaced the noisy "Library — Per Day" section on the Stats → Trends page with a "Library — Summary" section. The new section shows a top-10 watch-time leaderboard and a sortable per-title table (watch time, videos, sessions, cards, words, lookups, lookups/100w, date range), all scoped to the current date range selector.

If you are uncertain about the format, copy the most recent fragment's structure exactly and replace only the body text and category.

- [ ] **Step 3: Run the default handoff gate**

Run the commands in sequence (stop and fix if any fails):

```bash
bun run typecheck
bun run test:fast
bun run test:env
bun run changelog:lint
```

Expected: all PASS.

- [ ] **Step 4: Run the full build + smoke test**

```bash
bun run build
bun run test:smoke:dist
```

Expected: PASS.

- [ ] **Step 5: Commit the fragment**

```bash
git add changes/stats-library-summary.md
git commit -m "docs(changelog): summarize library summary replacing per-day trends"
```

---

## Verification Checklist

After all tasks complete, manually verify:

- [ ] The "Library — Per Day" section is gone from the Trends tab.
- [ ] A new "Library — Summary" section appears with a top-10 watch-time bar chart above a per-title table.
- [ ] Clicking table column headers sorts the table; clicking twice reverses direction.
- [ ] The shared Anime Visibility filter still hides titles from both the leaderboard, the table, and the Cumulative section below.
- [ ] Changing the date range selector (7d/30d/90d/365d/all) updates the summary.
- [ ] Titles with `words === 0` show `—` in the Lookups/100w column.
- [ ] Empty window shows "No library activity in the selected window."
- [ ] The "Library — Cumulative" section below is unchanged.
- [ ] `git log --oneline` shows small, focused commits per task.
