# Stats Dashboard v2 Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Redesign the stats dashboard to focus on session/media history with an activity feed, cover art library, and per-anime drill-down — while fixing the watch time inflation bug and relative date formatting.

**Architecture:** Activity feed as the default Overview tab, dedicated Library tab with Anilist cover art grid, per-anime detail view navigated from library cards. Bug fixes first, then backend (queries, API, rate limiter), then frontend (tabs, components, hooks).

**Tech Stack:** React 19, Recharts, Tailwind CSS (Catppuccin Macchiato), Hono server, SQLite, Anilist GraphQL API, Electron IPC

---

### Task 1: Fix Watch Time Inflation — Session Summaries Query

**Files:**
- Modify: `src/core/services/immersion-tracker/query.ts:11-34`

**Step 1: Fix `getSessionSummaries` to use MAX instead of SUM**

The telemetry values are cumulative snapshots. Each row stores the running total. Using `SUM()` across all telemetry rows for a session inflates values massively. Since the query already groups by `s.session_id`, change every `SUM(t.*)` to `MAX(t.*)`:

```typescript
export function getSessionSummaries(db: DatabaseSync, limit = 50): SessionSummaryQueryRow[] {
  const prepared = db.prepare(`
    SELECT
      s.session_id AS sessionId,
      s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      s.started_at_ms AS startedAtMs,
      s.ended_at_ms AS endedAtMs,
      COALESCE(MAX(t.total_watched_ms), 0) AS totalWatchedMs,
      COALESCE(MAX(t.active_watched_ms), 0) AS activeWatchedMs,
      COALESCE(MAX(t.lines_seen), 0) AS linesSeen,
      COALESCE(MAX(t.words_seen), 0) AS wordsSeen,
      COALESCE(MAX(t.tokens_seen), 0) AS tokensSeen,
      COALESCE(MAX(t.cards_mined), 0) AS cardsMined,
      COALESCE(MAX(t.lookup_count), 0) AS lookupCount,
      COALESCE(MAX(t.lookup_hits), 0) AS lookupHits
    FROM imm_sessions s
    LEFT JOIN imm_session_telemetry t ON t.session_id = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    GROUP BY s.session_id
    ORDER BY s.started_at_ms DESC
    LIMIT ?
  `);
  return prepared.all(limit) as unknown as SessionSummaryQueryRow[];
}
```

**Step 2: Verify build compiles**

Run: `cd /home/sudacode/projects/japanese/SubMiner && npx tsc --noEmit`

**Step 3: Commit**

```bash
git add src/core/services/immersion-tracker/query.ts
git commit -m "fix(stats): use MAX instead of SUM for cumulative telemetry in session summaries"
```

---

### Task 2: Fix Watch Time Inflation — Daily & Monthly Rollups

**Files:**
- Modify: `src/core/services/immersion-tracker/maintenance.ts:99-208`

**Step 1: Fix `upsertDailyRollupsForGroups` to use MAX-per-session subquery**

The rollup query must first get `MAX()` per session, then `SUM()` across sessions for that day+video combo:

```typescript
function upsertDailyRollupsForGroups(
  db: DatabaseSync,
  groups: Array<{ rollupDay: number; videoId: number }>,
  rollupNowMs: number,
): void {
  if (groups.length === 0) {
    return;
  }

  const upsertStmt = db.prepare(`
    INSERT INTO imm_daily_rollups (
      rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
      total_words_seen, total_tokens_seen, total_cards, cards_per_hour,
      words_per_min, lookup_hit_rate, CREATED_DATE, LAST_UPDATE_DATE
    )
    SELECT
      CAST(s.started_at_ms / 86400000 AS INTEGER) AS rollup_day,
      s.video_id AS video_id,
      COUNT(DISTINCT s.session_id) AS total_sessions,
      COALESCE(SUM(sm.max_active_ms), 0) / 60000.0 AS total_active_min,
      COALESCE(SUM(sm.max_lines), 0) AS total_lines_seen,
      COALESCE(SUM(sm.max_words), 0) AS total_words_seen,
      COALESCE(SUM(sm.max_tokens), 0) AS total_tokens_seen,
      COALESCE(SUM(sm.max_cards), 0) AS total_cards,
      CASE
        WHEN COALESCE(SUM(sm.max_active_ms), 0) > 0
          THEN (COALESCE(SUM(sm.max_cards), 0) * 60.0) / (COALESCE(SUM(sm.max_active_ms), 0) / 60000.0)
        ELSE NULL
      END AS cards_per_hour,
      CASE
        WHEN COALESCE(SUM(sm.max_active_ms), 0) > 0
          THEN COALESCE(SUM(sm.max_words), 0) / (COALESCE(SUM(sm.max_active_ms), 0) / 60000.0)
        ELSE NULL
      END AS words_per_min,
      CASE
        WHEN COALESCE(SUM(sm.max_lookups), 0) > 0
          THEN CAST(COALESCE(SUM(sm.max_hits), 0) AS REAL) / CAST(SUM(sm.max_lookups) AS REAL)
        ELSE NULL
      END AS lookup_hit_rate,
      ? AS CREATED_DATE,
      ? AS LAST_UPDATE_DATE
    FROM (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.lines_seen) AS max_lines,
        MAX(t.words_seen) AS max_words,
        MAX(t.tokens_seen) AS max_tokens,
        MAX(t.cards_mined) AS max_cards,
        MAX(t.lookup_count) AS max_lookups,
        MAX(t.lookup_hits) AS max_hits
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm
    JOIN imm_sessions s ON s.session_id = sm.session_id
    WHERE CAST(s.started_at_ms / 86400000 AS INTEGER) = ? AND s.video_id = ?
    GROUP BY rollup_day, s.video_id
    ON CONFLICT (rollup_day, video_id) DO UPDATE SET
      total_sessions = excluded.total_sessions,
      total_active_min = excluded.total_active_min,
      total_lines_seen = excluded.total_lines_seen,
      total_words_seen = excluded.total_words_seen,
      total_tokens_seen = excluded.total_tokens_seen,
      total_cards = excluded.total_cards,
      cards_per_hour = excluded.cards_per_hour,
      words_per_min = excluded.words_per_min,
      lookup_hit_rate = excluded.lookup_hit_rate,
      CREATED_DATE = COALESCE(imm_daily_rollups.CREATED_DATE, excluded.CREATED_DATE),
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `);

  for (const { rollupDay, videoId } of groups) {
    upsertStmt.run(rollupNowMs, rollupNowMs, rollupDay, videoId);
  }
}
```

**Step 2: Apply the same fix to `upsertMonthlyRollupsForGroups`**

Same subquery pattern — replace the direct `SUM(t.*)` with `SUM(sm.max_*)` via a `MAX`-per-session subquery:

```typescript
function upsertMonthlyRollupsForGroups(
  db: DatabaseSync,
  groups: Array<{ rollupMonth: number; videoId: number }>,
  rollupNowMs: number,
): void {
  if (groups.length === 0) {
    return;
  }

  const upsertStmt = db.prepare(`
    INSERT INTO imm_monthly_rollups (
      rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
      total_words_seen, total_tokens_seen, total_cards, CREATED_DATE, LAST_UPDATE_DATE
    )
    SELECT
      CAST(strftime('%Y%m', s.started_at_ms / 1000, 'unixepoch') AS INTEGER) AS rollup_month,
      s.video_id AS video_id,
      COUNT(DISTINCT s.session_id) AS total_sessions,
      COALESCE(SUM(sm.max_active_ms), 0) / 60000.0 AS total_active_min,
      COALESCE(SUM(sm.max_lines), 0) AS total_lines_seen,
      COALESCE(SUM(sm.max_words), 0) AS total_words_seen,
      COALESCE(SUM(sm.max_tokens), 0) AS total_tokens_seen,
      COALESCE(SUM(sm.max_cards), 0) AS total_cards,
      ? AS CREATED_DATE,
      ? AS LAST_UPDATE_DATE
    FROM (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.lines_seen) AS max_lines,
        MAX(t.words_seen) AS max_words,
        MAX(t.tokens_seen) AS max_tokens,
        MAX(t.cards_mined) AS max_cards
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm
    JOIN imm_sessions s ON s.session_id = sm.session_id
    WHERE CAST(strftime('%Y%m', s.started_at_ms / 1000, 'unixepoch') AS INTEGER) = ? AND s.video_id = ?
    GROUP BY rollup_month, s.video_id
    ON CONFLICT (rollup_month, video_id) DO UPDATE SET
      total_sessions = excluded.total_sessions,
      total_active_min = excluded.total_active_min,
      total_lines_seen = excluded.total_lines_seen,
      total_words_seen = excluded.total_words_seen,
      total_tokens_seen = excluded.total_tokens_seen,
      total_cards = excluded.total_cards,
      CREATED_DATE = COALESCE(imm_monthly_rollups.CREATED_DATE, excluded.CREATED_DATE),
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `);

  for (const { rollupMonth, videoId } of groups) {
    upsertStmt.run(rollupNowMs, rollupNowMs, rollupMonth, videoId);
  }
}
```

**Step 3: Verify build**

Run: `cd /home/sudacode/projects/japanese/SubMiner && npx tsc --noEmit`

**Step 4: Commit**

```bash
git add src/core/services/immersion-tracker/maintenance.ts
git commit -m "fix(stats): use MAX-per-session subquery in daily and monthly rollup aggregation"
```

---

### Task 3: Force-Rebuild Rollups on Schema Upgrade

**Files:**
- Modify: `src/core/services/immersion-tracker/storage.ts`
- Modify: `src/core/services/immersion-tracker/types.ts:1`

**Step 1: Bump schema version to trigger rebuild**

In `types.ts`, change line 1:
```typescript
export const SCHEMA_VERSION = 4;
```

**Step 2: Add rollup rebuild to schema migration in `storage.ts`**

At the end of `ensureSchema()`, before the `INSERT INTO imm_schema_version`, add a rollup wipe so that `runRollupMaintenance(db, true)` will recompute from scratch on next maintenance run:

```typescript
  // Wipe stale rollups so they get recomputed with corrected MAX-per-session logic
  if (currentVersion?.schema_version && currentVersion.schema_version < SCHEMA_VERSION) {
    db.exec('DELETE FROM imm_daily_rollups');
    db.exec('DELETE FROM imm_monthly_rollups');
    db.exec(`UPDATE imm_rollup_state SET state_value = 0 WHERE state_key = 'last_rollup_sample_ms'`);
  }
```

Add this block just before the final `INSERT INTO imm_schema_version` statement (before line 302).

**Step 3: Verify build**

Run: `cd /home/sudacode/projects/japanese/SubMiner && npx tsc --noEmit`

**Step 4: Commit**

```bash
git add src/core/services/immersion-tracker/types.ts src/core/services/immersion-tracker/storage.ts
git commit -m "fix(stats): bump schema to v4 and wipe rollups for recomputation"
```

---

### Task 4: Fix Relative Date Formatting

**Files:**
- Modify: `stats/src/lib/formatters.ts:18-26`
- Modify: `stats/src/lib/formatters.test.ts`

**Step 1: Update tests first**

Replace `stats/src/lib/formatters.test.ts` with comprehensive tests:

```typescript
import assert from 'node:assert/strict';
import test from 'node:test';

import { formatRelativeDate } from './formatters';

test('formatRelativeDate: future timestamps return "just now"', () => {
  assert.equal(formatRelativeDate(Date.now() + 60_000), 'just now');
});

test('formatRelativeDate: 0ms ago returns "just now"', () => {
  assert.equal(formatRelativeDate(Date.now()), 'just now');
});

test('formatRelativeDate: 30s ago returns "just now"', () => {
  assert.equal(formatRelativeDate(Date.now() - 30_000), 'just now');
});

test('formatRelativeDate: 5 minutes ago returns "5m ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 5 * 60_000), '5m ago');
});

test('formatRelativeDate: 59 minutes ago returns "59m ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 59 * 60_000), '59m ago');
});

test('formatRelativeDate: 2 hours ago returns "2h ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 2 * 3_600_000), '2h ago');
});

test('formatRelativeDate: 23 hours ago returns "23h ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 23 * 3_600_000), '23h ago');
});

test('formatRelativeDate: 36 hours ago returns "Yesterday"', () => {
  assert.equal(formatRelativeDate(Date.now() - 36 * 3_600_000), 'Yesterday');
});

test('formatRelativeDate: 5 days ago returns "5d ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 5 * 86_400_000), '5d ago');
});

test('formatRelativeDate: 10 days ago returns locale date string', () => {
  const ts = Date.now() - 10 * 86_400_000;
  assert.equal(formatRelativeDate(ts), new Date(ts).toLocaleDateString());
});
```

**Step 2: Run tests to verify they fail**

Run: `cd /home/sudacode/projects/japanese/SubMiner/stats && bun test src/lib/formatters.test.ts`
Expected: Several failures (current implementation lacks minute/hour granularity)

**Step 3: Implement the new formatter**

Replace `formatRelativeDate` in `stats/src/lib/formatters.ts`:

```typescript
export function formatRelativeDate(ms: number): string {
  const now = Date.now();
  const diffMs = now - ms;
  if (diffMs < 60_000) return 'just now';
  const diffMin = Math.floor(diffMs / 60_000);
  if (diffMin < 60) return `${diffMin}m ago`;
  const diffHours = Math.floor(diffMs / 3_600_000);
  if (diffHours < 24) return `${diffHours}h ago`;
  const diffDays = Math.floor(diffMs / 86_400_000);
  if (diffDays < 2) return 'Yesterday';
  if (diffDays < 7) return `${diffDays}d ago`;
  return new Date(ms).toLocaleDateString();
}
```

**Step 4: Run tests to verify they pass**

Run: `cd /home/sudacode/projects/japanese/SubMiner/stats && bun test src/lib/formatters.test.ts`
Expected: All pass

**Step 5: Commit**

```bash
git add stats/src/lib/formatters.ts stats/src/lib/formatters.test.ts
git commit -m "fix(stats): add minute and hour granularity to relative date formatting"
```

---

### Task 5: Add `imm_media_art` Table and Cover Art Queries

**Files:**
- Modify: `src/core/services/immersion-tracker/storage.ts` (add table in `ensureSchema`)
- Modify: `src/core/services/immersion-tracker/query.ts` (add new query functions)
- Modify: `src/core/services/immersion-tracker/types.ts` (add new row types)

**Step 1: Add types**

Append to `src/core/services/immersion-tracker/types.ts`:

```typescript
export interface MediaArtRow {
  videoId: number;
  anilistId: number | null;
  coverUrl: string | null;
  coverBlob: Buffer | null;
  titleRomaji: string | null;
  titleEnglish: string | null;
  episodesTotal: number | null;
  fetchedAtMs: number;
}

export interface MediaLibraryRow {
  videoId: number;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  lastWatchedMs: number;
  hasCoverArt: number;
}

export interface MediaDetailRow {
  videoId: number;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  totalLinesSeen: number;
  totalLookupCount: number;
  totalLookupHits: number;
}
```

**Step 2: Add table creation in `ensureSchema`**

Add after the `imm_kanji` table creation block (after line 191 in storage.ts):

```typescript
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_media_art(
      video_id INTEGER PRIMARY KEY,
      anilist_id INTEGER,
      cover_url TEXT,
      cover_blob BLOB,
      title_romaji TEXT,
      title_english TEXT,
      episodes_total INTEGER,
      fetched_at_ms INTEGER NOT NULL,
      CREATED_DATE INTEGER,
      LAST_UPDATE_DATE INTEGER,
      FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE
    );
  `);
```

**Step 3: Add query functions**

Append to `src/core/services/immersion-tracker/query.ts`:

```typescript
import type { MediaArtRow, MediaLibraryRow, MediaDetailRow } from './types';

export function getMediaLibrary(db: DatabaseSync): MediaLibraryRow[] {
  return db.prepare(`
    SELECT
      v.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      COUNT(DISTINCT s.session_id) AS totalSessions,
      COALESCE(SUM(sm.max_active_ms), 0) AS totalActiveMs,
      COALESCE(SUM(sm.max_cards), 0) AS totalCards,
      COALESCE(SUM(sm.max_words), 0) AS totalWordsSeen,
      MAX(s.started_at_ms) AS lastWatchedMs,
      CASE WHEN ma.cover_blob IS NOT NULL THEN 1 ELSE 0 END AS hasCoverArt
    FROM imm_videos v
    JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.cards_mined) AS max_cards,
        MAX(t.words_seen) AS max_words
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm ON sm.session_id = s.session_id
    LEFT JOIN imm_media_art ma ON ma.video_id = v.video_id
    GROUP BY v.video_id
    ORDER BY lastWatchedMs DESC
  `).all() as unknown as MediaLibraryRow[];
}

export function getMediaDetail(db: DatabaseSync, videoId: number): MediaDetailRow | null {
  return db.prepare(`
    SELECT
      v.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      COUNT(DISTINCT s.session_id) AS totalSessions,
      COALESCE(SUM(sm.max_active_ms), 0) AS totalActiveMs,
      COALESCE(SUM(sm.max_cards), 0) AS totalCards,
      COALESCE(SUM(sm.max_words), 0) AS totalWordsSeen,
      COALESCE(SUM(sm.max_lines), 0) AS totalLinesSeen,
      COALESCE(SUM(sm.max_lookups), 0) AS totalLookupCount,
      COALESCE(SUM(sm.max_hits), 0) AS totalLookupHits
    FROM imm_videos v
    JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.cards_mined) AS max_cards,
        MAX(t.words_seen) AS max_words,
        MAX(t.lines_seen) AS max_lines,
        MAX(t.lookup_count) AS max_lookups,
        MAX(t.lookup_hits) AS max_hits
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm ON sm.session_id = s.session_id
    WHERE v.video_id = ?
    GROUP BY v.video_id
  `).get(videoId) as unknown as MediaDetailRow | null;
}

export function getMediaSessions(db: DatabaseSync, videoId: number, limit = 100): SessionSummaryQueryRow[] {
  return db.prepare(`
    SELECT
      s.session_id AS sessionId,
      s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      s.started_at_ms AS startedAtMs,
      s.ended_at_ms AS endedAtMs,
      COALESCE(MAX(t.total_watched_ms), 0) AS totalWatchedMs,
      COALESCE(MAX(t.active_watched_ms), 0) AS activeWatchedMs,
      COALESCE(MAX(t.lines_seen), 0) AS linesSeen,
      COALESCE(MAX(t.words_seen), 0) AS wordsSeen,
      COALESCE(MAX(t.tokens_seen), 0) AS tokensSeen,
      COALESCE(MAX(t.cards_mined), 0) AS cardsMined,
      COALESCE(MAX(t.lookup_count), 0) AS lookupCount,
      COALESCE(MAX(t.lookup_hits), 0) AS lookupHits
    FROM imm_sessions s
    LEFT JOIN imm_session_telemetry t ON t.session_id = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    WHERE s.video_id = ?
    GROUP BY s.session_id
    ORDER BY s.started_at_ms DESC
    LIMIT ?
  `).all(videoId, limit) as unknown as SessionSummaryQueryRow[];
}

export function getMediaDailyRollups(db: DatabaseSync, videoId: number, limit = 90): ImmersionSessionRollupRow[] {
  return db.prepare(`
    SELECT
      rollup_day AS rollupDayOrMonth,
      video_id AS videoId,
      total_sessions AS totalSessions,
      total_active_min AS totalActiveMin,
      total_lines_seen AS totalLinesSeen,
      total_words_seen AS totalWordsSeen,
      total_tokens_seen AS totalTokensSeen,
      total_cards AS totalCards,
      cards_per_hour AS cardsPerHour,
      words_per_min AS wordsPerMin,
      lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups
    WHERE video_id = ?
    ORDER BY rollup_day DESC
    LIMIT ?
  `).all(videoId, limit) as unknown as ImmersionSessionRollupRow[];
}

export function getCoverArt(db: DatabaseSync, videoId: number): MediaArtRow | null {
  return db.prepare(`
    SELECT
      video_id AS videoId,
      anilist_id AS anilistId,
      cover_url AS coverUrl,
      cover_blob AS coverBlob,
      title_romaji AS titleRomaji,
      title_english AS titleEnglish,
      episodes_total AS episodesTotal,
      fetched_at_ms AS fetchedAtMs
    FROM imm_media_art
    WHERE video_id = ?
  `).get(videoId) as unknown as MediaArtRow | null;
}

export function upsertCoverArt(
  db: DatabaseSync,
  videoId: number,
  art: {
    anilistId: number | null;
    coverUrl: string | null;
    coverBlob: Buffer | null;
    titleRomaji: string | null;
    titleEnglish: string | null;
    episodesTotal: number | null;
  },
): void {
  const nowMs = Date.now();
  db.prepare(`
    INSERT INTO imm_media_art (
      video_id, anilist_id, cover_url, cover_blob,
      title_romaji, title_english, episodes_total,
      fetched_at_ms, CREATED_DATE, LAST_UPDATE_DATE
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(video_id) DO UPDATE SET
      anilist_id = excluded.anilist_id,
      cover_url = excluded.cover_url,
      cover_blob = excluded.cover_blob,
      title_romaji = excluded.title_romaji,
      title_english = excluded.title_english,
      episodes_total = excluded.episodes_total,
      fetched_at_ms = excluded.fetched_at_ms,
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `).run(
    videoId, art.anilistId, art.coverUrl, art.coverBlob,
    art.titleRomaji, art.titleEnglish, art.episodesTotal,
    nowMs, nowMs, nowMs,
  );
}
```

**Step 4: Verify build**

Run: `cd /home/sudacode/projects/japanese/SubMiner && npx tsc --noEmit`

**Step 5: Commit**

```bash
git add src/core/services/immersion-tracker/types.ts src/core/services/immersion-tracker/storage.ts src/core/services/immersion-tracker/query.ts
git commit -m "feat(stats): add imm_media_art table and media library/detail queries"
```

---

### Task 6: Centralized Anilist Rate Limiter

**Files:**
- Create: `src/core/services/anilist/rate-limiter.ts`

**Step 1: Implement sliding-window rate limiter**

```typescript
const DEFAULT_MAX_PER_MINUTE = 20;
const WINDOW_MS = 60_000;
const SAFETY_REMAINING_THRESHOLD = 5;

export interface AnilistRateLimiter {
  acquire(): Promise<void>;
  recordResponse(headers: Headers): void;
}

export function createAnilistRateLimiter(
  maxPerMinute = DEFAULT_MAX_PER_MINUTE,
): AnilistRateLimiter {
  const timestamps: number[] = [];
  let pauseUntilMs = 0;

  function pruneOld(now: number): void {
    const cutoff = now - WINDOW_MS;
    while (timestamps.length > 0 && timestamps[0]! < cutoff) {
      timestamps.shift();
    }
  }

  return {
    async acquire(): Promise<void> {
      const now = Date.now();

      if (now < pauseUntilMs) {
        const waitMs = pauseUntilMs - now;
        await new Promise((resolve) => setTimeout(resolve, waitMs));
      }

      pruneOld(Date.now());

      if (timestamps.length >= maxPerMinute) {
        const oldest = timestamps[0]!;
        const waitMs = oldest + WINDOW_MS - Date.now() + 100;
        if (waitMs > 0) {
          await new Promise((resolve) => setTimeout(resolve, waitMs));
        }
        pruneOld(Date.now());
      }

      timestamps.push(Date.now());
    },

    recordResponse(headers: Headers): void {
      const remaining = headers.get('x-ratelimit-remaining');
      if (remaining !== null) {
        const n = parseInt(remaining, 10);
        if (Number.isFinite(n) && n < SAFETY_REMAINING_THRESHOLD) {
          const reset = headers.get('x-ratelimit-reset');
          if (reset) {
            const resetMs = parseInt(reset, 10) * 1000;
            if (Number.isFinite(resetMs)) {
              pauseUntilMs = Math.max(pauseUntilMs, resetMs);
            }
          } else {
            pauseUntilMs = Math.max(pauseUntilMs, Date.now() + WINDOW_MS);
          }
        }
      }

      const retryAfter = headers.get('retry-after');
      if (retryAfter) {
        const seconds = parseInt(retryAfter, 10);
        if (Number.isFinite(seconds) && seconds > 0) {
          pauseUntilMs = Math.max(pauseUntilMs, Date.now() + seconds * 1000);
        }
      }
    },
  };
}
```

**Step 2: Verify build**

Run: `cd /home/sudacode/projects/japanese/SubMiner && npx tsc --noEmit`

**Step 3: Commit**

```bash
git add src/core/services/anilist/rate-limiter.ts
git commit -m "feat(stats): add centralized Anilist rate limiter with sliding window"
```

---

### Task 7: Cover Art Fetcher Service

**Files:**
- Create: `src/core/services/anilist/cover-art-fetcher.ts`

**Step 1: Implement the cover art fetcher**

This service searches Anilist for anime cover art and caches results. It reuses the existing `guessAnilistMediaInfo` for title parsing and `pickBestSearchResult`-style matching.

```typescript
import type { DatabaseSync } from '../immersion-tracker/sqlite';
import type { AnilistRateLimiter } from './rate-limiter';
import { getCoverArt, upsertCoverArt } from '../immersion-tracker/query';

const ANILIST_GRAPHQL_URL = 'https://graphql.anilist.co';

const SEARCH_QUERY = `
  query ($search: String!) {
    Page(perPage: 5) {
      media(search: $search, type: ANIME) {
        id
        episodes
        coverImage { large medium }
        title { romaji english native }
      }
    }
  }
`;

interface AnilistSearchMedia {
  id: number;
  episodes: number | null;
  coverImage?: { large?: string; medium?: string };
  title?: { romaji?: string; english?: string; native?: string };
}

interface AnilistSearchResponse {
  data?: { Page?: { media?: AnilistSearchMedia[] } };
  errors?: Array<{ message?: string }>;
}

function stripFilenameTags(title: string): string {
  return title
    .replace(/\s*\[.*?\]\s*/g, ' ')
    .replace(/\s*\((?:\d{4}|(?:\d+(?:bit|p)))\)\s*/gi, ' ')
    .replace(/\s*-\s*S\d+E\d+\s*/i, ' ')
    .replace(/\s*-\s*\d{2,4}\s*/, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export interface CoverArtFetcher {
  fetchIfMissing(db: DatabaseSync, videoId: number, canonicalTitle: string): Promise<boolean>;
}

export function createCoverArtFetcher(
  rateLimiter: AnilistRateLimiter,
  logger: { info: (msg: string) => void; warn: (msg: string, detail?: unknown) => void },
): CoverArtFetcher {
  return {
    async fetchIfMissing(db: DatabaseSync, videoId: number, canonicalTitle: string): Promise<boolean> {
      const existing = getCoverArt(db, videoId);
      if (existing) return true;

      const searchTitle = stripFilenameTags(canonicalTitle);
      if (!searchTitle) {
        upsertCoverArt(db, videoId, {
          anilistId: null, coverUrl: null, coverBlob: null,
          titleRomaji: null, titleEnglish: null, episodesTotal: null,
        });
        return false;
      }

      try {
        await rateLimiter.acquire();
        const res = await fetch(ANILIST_GRAPHQL_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ query: SEARCH_QUERY, variables: { search: searchTitle } }),
        });
        rateLimiter.recordResponse(res.headers);

        if (res.status === 429) {
          logger.warn(`Anilist 429 for "${searchTitle}", will retry later`);
          return false;
        }

        const payload = await res.json() as AnilistSearchResponse;
        const media = payload.data?.Page?.media ?? [];
        if (media.length === 0) {
          upsertCoverArt(db, videoId, {
            anilistId: null, coverUrl: null, coverBlob: null,
            titleRomaji: null, titleEnglish: null, episodesTotal: null,
          });
          return false;
        }

        const best = media[0]!;
        const coverUrl = best.coverImage?.large ?? best.coverImage?.medium ?? null;
        let coverBlob: Buffer | null = null;

        if (coverUrl) {
          await rateLimiter.acquire();
          const imgRes = await fetch(coverUrl);
          rateLimiter.recordResponse(imgRes.headers);
          if (imgRes.ok) {
            coverBlob = Buffer.from(await imgRes.arrayBuffer());
          }
        }

        upsertCoverArt(db, videoId, {
          anilistId: best.id,
          coverUrl,
          coverBlob,
          titleRomaji: best.title?.romaji ?? null,
          titleEnglish: best.title?.english ?? null,
          episodesTotal: best.episodes,
        });

        logger.info(`Cached cover art for "${searchTitle}" (anilist:${best.id})`);
        return true;
      } catch (err) {
        logger.warn(`Cover art fetch failed for "${searchTitle}"`, err);
        return false;
      }
    },
  };
}
```

**Step 2: Verify build**

Run: `cd /home/sudacode/projects/japanese/SubMiner && npx tsc --noEmit`

**Step 3: Commit**

```bash
git add src/core/services/anilist/cover-art-fetcher.ts
git commit -m "feat(stats): add cover art fetcher with Anilist search and image caching"
```

---

### Task 8: Add Media API Endpoints and IPC Handlers

**Files:**
- Modify: `src/core/services/stats-server.ts`
- Modify: `src/core/services/ipc.ts`
- Modify: `src/shared/ipc/contracts.ts`
- Modify: `src/preload-stats.ts`

**Step 1: Add new IPC channel constants**

In `src/shared/ipc/contracts.ts`, add to `IPC_CHANNELS.request` (after line 72):

```typescript
    statsGetMediaLibrary: 'stats:get-media-library',
    statsGetMediaDetail: 'stats:get-media-detail',
    statsGetMediaSessions: 'stats:get-media-sessions',
    statsGetMediaDailyRollups: 'stats:get-media-daily-rollups',
    statsGetMediaCover: 'stats:get-media-cover',
```

**Step 2: Add HTTP routes to stats-server.ts**

Add before the `return app;` line in `createStatsApp()`:

```typescript
  app.get('/api/stats/media', async (c) => {
    const library = await tracker.getMediaLibrary();
    return c.json(library);
  });

  app.get('/api/stats/media/:videoId', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.json(null, 400);
    const [detail, sessions, rollups] = await Promise.all([
      tracker.getMediaDetail(videoId),
      tracker.getMediaSessions(videoId, 100),
      tracker.getMediaDailyRollups(videoId, 90),
    ]);
    return c.json({ detail, sessions, rollups });
  });

  app.get('/api/stats/media/:videoId/cover', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.body(null, 404);
    const art = await tracker.getCoverArt(videoId);
    if (!art?.coverBlob) return c.body(null, 404);
    return new Response(art.coverBlob, {
      headers: {
        'Content-Type': 'image/jpeg',
        'Cache-Control': 'public, max-age=604800',
      },
    });
  });
```

**Step 3: Add IPC handlers**

Add corresponding IPC handlers in `src/core/services/ipc.ts` following the existing pattern (after the `statsGetKanji` handler).

**Step 4: Add preload API methods**

Add to `src/preload-stats.ts` statsAPI object:

```typescript
  getMediaLibrary: (): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaLibrary),

  getMediaDetail: (videoId: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaDetail, videoId),

  getMediaSessions: (videoId: number, limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaSessions, videoId, limit),

  getMediaDailyRollups: (videoId: number, limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaDailyRollups, videoId, limit),

  getMediaCover: (videoId: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaCover, videoId),
```

**Step 5: Wire up `ImmersionTrackerService` to expose the new query methods**

The service needs to expose `getMediaLibrary()`, `getMediaDetail(videoId)`, `getMediaSessions(videoId, limit)`, `getMediaDailyRollups(videoId, limit)`, and `getCoverArt(videoId)` by delegating to the query functions added in Task 5.

**Step 6: Verify build**

Run: `cd /home/sudacode/projects/japanese/SubMiner && npx tsc --noEmit`

**Step 7: Commit**

```bash
git add src/shared/ipc/contracts.ts src/core/services/stats-server.ts src/core/services/ipc.ts src/preload-stats.ts src/core/services/immersion-tracker-service.ts
git commit -m "feat(stats): add media library/detail/cover API endpoints and IPC handlers"
```

---

### Task 9: Frontend — Update Types, Clients, and Hooks

**Files:**
- Modify: `stats/src/types/stats.ts`
- Modify: `stats/src/lib/api-client.ts`
- Modify: `stats/src/lib/ipc-client.ts`
- Create: `stats/src/hooks/useMediaLibrary.ts`
- Create: `stats/src/hooks/useMediaDetail.ts`

**Step 1: Add new types in `stats/src/types/stats.ts`**

```typescript
export interface MediaLibraryItem {
  videoId: number;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  lastWatchedMs: number;
  hasCoverArt: number;
}

export interface MediaDetailData {
  detail: {
    videoId: number;
    canonicalTitle: string;
    totalSessions: number;
    totalActiveMs: number;
    totalCards: number;
    totalWordsSeen: number;
    totalLinesSeen: number;
    totalLookupCount: number;
    totalLookupHits: number;
  } | null;
  sessions: SessionSummary[];
  rollups: DailyRollup[];
}
```

**Step 2: Add new methods to both clients**

Add to `apiClient` in `stats/src/lib/api-client.ts`:

```typescript
  getMediaLibrary: () => fetchJson<MediaLibraryItem[]>('/api/stats/media'),
  getMediaDetail: (videoId: number) =>
    fetchJson<MediaDetailData>(`/api/stats/media/${videoId}`),
```

Add matching methods to `ipcClient` in `stats/src/lib/ipc-client.ts` and the `StatsElectronAPI` interface.

**Step 3: Create `stats/src/hooks/useMediaLibrary.ts`**

```typescript
import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { MediaLibraryItem } from '../types/stats';

export function useMediaLibrary() {
  const [media, setMedia] = useState<MediaLibraryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    getStatsClient()
      .getMediaLibrary()
      .then(setMedia)
      .catch((err: Error) => setError(err.message))
      .finally(() => setLoading(false));
  }, []);

  return { media, loading, error };
}
```

**Step 4: Create `stats/src/hooks/useMediaDetail.ts`**

```typescript
import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { MediaDetailData } from '../types/stats';

export function useMediaDetail(videoId: number | null) {
  const [data, setData] = useState<MediaDetailData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (videoId === null) return;
    setLoading(true);
    setError(null);
    getStatsClient()
      .getMediaDetail(videoId)
      .then(setData)
      .catch((err: Error) => setError(err.message))
      .finally(() => setLoading(false));
  }, [videoId]);

  return { data, loading, error };
}
```

**Step 5: Verify frontend build**

Run: `cd /home/sudacode/projects/japanese/SubMiner/stats && bun run build`

**Step 6: Commit**

```bash
git add stats/src/types/stats.ts stats/src/lib/api-client.ts stats/src/lib/ipc-client.ts stats/src/hooks/useMediaLibrary.ts stats/src/hooks/useMediaDetail.ts
git commit -m "feat(stats): add media library and detail types, clients, and hooks"
```

---

### Task 10: Frontend — Redesign Overview Tab as Activity Feed

**Files:**
- Modify: `stats/src/components/overview/OverviewTab.tsx`
- Modify: `stats/src/components/overview/HeroStats.tsx`
- Modify: `stats/src/components/overview/RecentSessions.tsx`
- Delete or repurpose: `stats/src/components/overview/QuickStats.tsx`

**Step 1: Simplify HeroStats to 4 cards: Watch Time Today, Cards Mined, Streak, All Time**

Replace the "Words Seen" and "Lookup Hit Rate" cards with "Streak" and "All Time" — move the streak logic from QuickStats into HeroStats. The all-time total is the sum of all rollup `totalActiveMin`.

**Step 2: Redesign RecentSessions as an activity feed**

- Group sessions by day ("Today", "Yesterday", "March 10")
- Each row: small cover art thumbnail (48x64), clean title, relative time + duration, cards + words stats
- Use the cover art endpoint: `/api/stats/media/${videoId}/cover` with an `<img>` tag and fallback placeholder

**Step 3: Remove WatchTimeChart and QuickStats from the Overview tab**

The watch time chart moves to the Trends tab. QuickStats data is absorbed into HeroStats.

**Step 4: Update OverviewTab layout**

```tsx
export function OverviewTab() {
  const { data, loading, error } = useOverview();
  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;
  if (!data) return null;

  return (
    <div className="space-y-4">
      <HeroStats data={data} />
      <RecentSessions sessions={data.sessions} />
    </div>
  );
}
```

**Step 5: Verify frontend build**

Run: `cd /home/sudacode/projects/japanese/SubMiner/stats && bun run build`

**Step 6: Commit**

```bash
git add stats/src/components/overview/
git commit -m "feat(stats): redesign Overview tab as activity feed with hero stats"
```

---

### Task 11: Frontend — Library Tab with Cover Art Grid

**Files:**
- Create: `stats/src/components/library/LibraryTab.tsx`
- Create: `stats/src/components/library/MediaCard.tsx`
- Create: `stats/src/components/library/CoverImage.tsx`

**Step 1: Create CoverImage component**

Loads cover art from `/api/stats/media/${videoId}/cover`. Falls back to a gray placeholder with the first character of the title. Handles loading state.

**Step 2: Create MediaCard component**

Shows: CoverImage (3:4 aspect ratio), episode badge, title, watch time, cards mined. Accepts `onClick` prop for navigation.

**Step 3: Create LibraryTab**

Uses `useMediaLibrary()` hook. Renders search input, filter chips (All/Watching/Completed — for v1, "All" only since we don't track watch status yet), summary line ("N titles · Xh total"), and a CSS grid of MediaCards. Clicking a card sets a `selectedVideoId` state to navigate to the detail view (Task 12).

**Step 4: Verify frontend build**

Run: `cd /home/sudacode/projects/japanese/SubMiner/stats && bun run build`

**Step 5: Commit**

```bash
git add stats/src/components/library/
git commit -m "feat(stats): add Library tab with cover art grid"
```

---

### Task 12: Frontend — Per-Anime Detail View

**Files:**
- Create: `stats/src/components/library/MediaDetailView.tsx`
- Create: `stats/src/components/library/MediaHeader.tsx`
- Create: `stats/src/components/library/MediaWatchChart.tsx`
- Create: `stats/src/components/library/MediaSessionList.tsx`

**Step 1: Create MediaHeader**

Large cover art on the left, title + stats on the right (total watch time, total episodes/sessions, cards mined, avg session length).

**Step 2: Create MediaWatchChart**

Reuse the existing `WatchTimeChart` pattern (Recharts BarChart) but scoped to the anime's rollups from `MediaDetailData.rollups`.

**Step 3: Create MediaSessionList**

List of sessions for this anime. Reuse the SessionRow pattern but without the expand/detail — just show timestamp, duration, cards, words per session.

**Step 4: Create MediaDetailView**

Composed component: back button, MediaHeader, MediaWatchChart, MediaSessionList. Uses `useMediaDetail(videoId)` hook. The vocabulary section can be a placeholder for now ("Coming soon") to keep v1 scope manageable.

**Step 5: Integrate into LibraryTab**

When `selectedVideoId` is set, render `MediaDetailView` instead of the grid. Back button resets `selectedVideoId` to null.

**Step 6: Verify frontend build**

Run: `cd /home/sudacode/projects/japanese/SubMiner/stats && bun run build`

**Step 7: Commit**

```bash
git add stats/src/components/library/
git commit -m "feat(stats): add per-anime detail view with header, chart, and session history"
```

---

### Task 13: Frontend — Update Tab Bar and App Shell

**Files:**
- Modify: `stats/src/components/layout/TabBar.tsx`
- Modify: `stats/src/App.tsx`

**Step 1: Update TabBar tabs**

Change `TabId` type and `TABS` array:

```typescript
export type TabId = 'overview' | 'library' | 'trends' | 'vocabulary';

const TABS: Tab[] = [
  { id: 'overview', label: 'Overview' },
  { id: 'library', label: 'Library' },
  { id: 'trends', label: 'Trends' },
  { id: 'vocabulary', label: 'Vocabulary' },
];
```

**Step 2: Update App.tsx**

Replace the Sessions tab panel with Library, import `LibraryTab`:

```tsx
import { LibraryTab } from './components/library/LibraryTab';

// In the JSX, replace the sessions section with:
<section id="panel-library" role="tabpanel" aria-labelledby="tab-library" hidden={activeTab !== 'library'}>
  <LibraryTab />
</section>
```

Remove the SessionsTab import. The Sessions tab functionality is now part of the activity feed (Overview) and per-anime detail (Library).

**Step 3: Verify frontend build**

Run: `cd /home/sudacode/projects/japanese/SubMiner/stats && bun run build`

**Step 4: Commit**

```bash
git add stats/src/components/layout/TabBar.tsx stats/src/App.tsx
git commit -m "feat(stats): replace Sessions tab with Library tab in app shell"
```

---

### Task 14: Integration Test — Full Build and Smoke Test

**Step 1: Full build**

Run: `cd /home/sudacode/projects/japanese/SubMiner && make build`

**Step 2: Run existing tests**

Run: `cd /home/sudacode/projects/japanese/SubMiner && bun test`

**Step 3: Manual smoke test**

Launch the app, open the stats overlay, verify:
- Overview tab shows activity feed with relative timestamps
- Watch time values are reasonable (not inflated)
- Library tab shows grid with cover art placeholders
- Clicking a card shows the detail view
- Back button returns to grid
- Trends and Vocabulary tabs still work

**Step 4: Final commit**

```bash
git add -A
git commit -m "feat(stats): stats dashboard v2 with activity feed, library grid, and per-anime detail"
```
