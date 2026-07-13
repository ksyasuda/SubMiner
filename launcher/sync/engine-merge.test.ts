import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Database } from 'bun:sqlite';
// The engine executes on libsql in production; these merge tests run through
// that same driver. bun:sqlite is used only to build fixtures and inspect
// results.
import { createDbSnapshot } from '../../src/core/services/stats-sync/shared.js';
import { mergeSnapshotIntoDb } from '../../src/core/services/stats-sync/merge.js';
import {
  createImmersionDbFixture,
  insertFixtureSession,
} from '../test-support/immersion-db-fixture.js';

const DAY_MS = 86_400_000;
const BASE_MS = Date.UTC(2026, 5, 1, 12, 0, 0);

function makeTmpDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-sync-test-'));
}

function makeDbPair(): { dir: string; localPath: string; remotePath: string } {
  const dir = makeTmpDir();
  const localPath = path.join(dir, 'local.sqlite');
  const remotePath = path.join(dir, 'remote.sqlite');
  createImmersionDbFixture(localPath);
  createImmersionDbFixture(remotePath);
  return { dir, localPath, remotePath };
}

function queryOne<T extends Record<string, unknown>>(
  dbPath: string,
  sql: string,
  params: unknown[] = [],
): T | undefined {
  const db = new Database(dbPath, { readonly: true });
  try {
    return db.query<T>(sql).get(...params) as T | undefined;
  } finally {
    db.close();
  }
}

function count(dbPath: string, sql: string, params: unknown[] = []): number {
  return Number(queryOne<{ n: number }>(dbPath, sql, params)?.n ?? 0);
}

function withWritableDb<T>(dbPath: string, fn: (db: Database) => T): T {
  const db = new Database(dbPath, { readwrite: true });
  try {
    return fn(db);
  } finally {
    db.close();
  }
}

test('merges remote-only sessions with catalog, lifetime, and rollups', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(localPath, {
      uuid: 'local-1',
      videoKey: 'showa-e1',
      animeTitleKey: 'showa',
      startedAtMs: BASE_MS,
      applyLifetime: true,
      words: [{ headword: '見る', word: '見た', reading: 'みた', count: 2 }],
    });
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'showb-e1',
      animeTitleKey: 'showb',
      startedAtMs: BASE_MS + DAY_MS,
      activeWatchedMs: 900_000,
      cardsMined: 3,
      applyLifetime: true,
      words: [
        { headword: '見る', word: '見た', reading: 'みた', count: 5 },
        { headword: '食べる', word: '食べた', reading: 'たべた', count: 1 },
      ],
    });

    const summary = mergeSnapshotIntoDb(localPath, remotePath);
    assert.equal(summary.sessionsMerged, 1);
    assert.equal(summary.animeAdded, 1);
    assert.equal(summary.videosAdded, 1);
    assert.equal(summary.wordsAdded, 1);
    assert.equal(summary.subtitleLinesAdded, 2);
    assert.equal(summary.telemetryRowsAdded, 1);

    assert.equal(count(localPath, 'SELECT COUNT(*) AS n FROM imm_sessions'), 2);
    const global = queryOne<{
      total_sessions: number;
      total_active_ms: number;
      total_cards: number;
      active_days: number;
      episodes_started: number;
    }>(
      localPath,
      'SELECT total_sessions, total_active_ms, total_cards, active_days, episodes_started FROM imm_lifetime_global WHERE global_id = 1',
    );
    assert.equal(global?.total_sessions, 2);
    assert.equal(global?.total_active_ms, 1_200_000 + 900_000);
    assert.equal(global?.total_cards, 2 + 3);
    assert.equal(global?.active_days, 2);
    assert.equal(global?.episodes_started, 2);

    // Existing word: local 2 + merged 5; new word carries remote frequency.
    const sharedWord = queryOne<{ frequency: number }>(
      localPath,
      `SELECT frequency FROM imm_words WHERE word = '見た'`,
    );
    assert.equal(sharedWord?.frequency, 7);
    const newWord = queryOne<{ frequency: number }>(
      localPath,
      `SELECT frequency FROM imm_words WHERE word = '食べた'`,
    );
    assert.equal(newWord?.frequency, 1);

    // The merged session's rollup group was recomputed.
    assert.equal(summary.rollupGroupsRecomputed, 1);
    const mergedVideoId = Number(
      queryOne<{ video_id: number }>(
        localPath,
        `SELECT video_id FROM imm_videos WHERE video_key = 'showb-e1'`,
      )?.video_id,
    );
    assert.equal(
      count(localPath, 'SELECT COUNT(*) AS n FROM imm_daily_rollups WHERE video_id = ?', [
        mergedVideoId,
      ]),
      1,
    );
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('is idempotent: re-merging the same snapshot changes nothing', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'showa-e1',
      animeTitleKey: 'showa',
      startedAtMs: BASE_MS,
      applyLifetime: true,
      words: [{ headword: '見る', word: '見た', reading: 'みた', count: 4 }],
    });

    mergeSnapshotIntoDb(localPath, remotePath);
    const globalAfterFirst = queryOne<Record<string, unknown>>(
      localPath,
      'SELECT * FROM imm_lifetime_global WHERE global_id = 1',
    );
    const summary = mergeSnapshotIntoDb(localPath, remotePath);

    assert.equal(summary.sessionsMerged, 0);
    assert.equal(summary.sessionsAlreadyPresent, 1);
    assert.equal(summary.wordsAdded, 0);
    assert.equal(count(localPath, 'SELECT COUNT(*) AS n FROM imm_sessions'), 1);
    assert.equal(count(localPath, 'SELECT COUNT(*) AS n FROM imm_subtitle_lines'), 1);
    const globalAfterSecond = queryOne<Record<string, unknown>>(
      localPath,
      'SELECT * FROM imm_lifetime_global WHERE global_id = 1',
    );
    assert.deepEqual(
      { ...globalAfterSecond, LAST_UPDATE_DATE: null },
      { ...globalAfterFirst, LAST_UPDATE_DATE: null },
    );
    assert.equal(
      Number(
        queryOne<{ frequency: number }>(
          localPath,
          `SELECT frequency FROM imm_words WHERE word = '見た'`,
        )?.frequency,
      ),
      4,
    );
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('preserves anilist_id when inserting a new anime from the snapshot', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'showb-e1',
      animeTitleKey: 'showb',
      startedAtMs: BASE_MS,
      applyLifetime: true,
    });
    withWritableDb(remotePath, (remoteDb) => {
      remoteDb
        .prepare(`UPDATE imm_anime SET anilist_id = 12345 WHERE normalized_title_key = 'showb'`)
        .run();
    });

    const summary = mergeSnapshotIntoDb(localPath, remotePath);
    assert.equal(summary.animeAdded, 1);
    const anime = queryOne<{ anilist_id: number }>(
      localPath,
      `SELECT anilist_id FROM imm_anime WHERE normalized_title_key = 'showb'`,
    );
    assert.equal(anime?.anilist_id, 12345);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('matches an existing local anime by anilist_id even when the title key differs', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(localPath, {
      uuid: 'local-1',
      videoKey: 'showa-e1',
      animeTitleKey: 'show-romaji',
      startedAtMs: BASE_MS,
      applyLifetime: true,
    });
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'showa-e2',
      animeTitleKey: 'show-native',
      startedAtMs: BASE_MS + DAY_MS,
      applyLifetime: true,
    });
    withWritableDb(localPath, (localDb) => {
      localDb
        .prepare(`UPDATE imm_anime SET anilist_id = 999 WHERE normalized_title_key = 'show-romaji'`)
        .run();
    });
    withWritableDb(remotePath, (remoteDb) => {
      remoteDb
        .prepare(`UPDATE imm_anime SET anilist_id = 999 WHERE normalized_title_key = 'show-native'`)
        .run();
    });

    const summary = mergeSnapshotIntoDb(localPath, remotePath);
    // Same anilist_id → one anime, not two.
    assert.equal(summary.animeAdded, 0);
    assert.equal(count(localPath, 'SELECT COUNT(*) AS n FROM imm_anime'), 1);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('matches shared videos by video_key and merges the watched flag', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(localPath, {
      uuid: 'local-1',
      videoKey: 'showa-e1',
      animeTitleKey: 'showa',
      startedAtMs: BASE_MS,
      applyLifetime: true,
    });
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'showa-e1',
      animeTitleKey: 'showa',
      startedAtMs: BASE_MS + DAY_MS,
      watched: true,
      applyLifetime: true,
    });

    const summary = mergeSnapshotIntoDb(localPath, remotePath);
    assert.equal(summary.videosAdded, 0);
    assert.equal(summary.animeAdded, 0);
    assert.equal(count(localPath, 'SELECT COUNT(*) AS n FROM imm_videos'), 1);
    const video = queryOne<{ watched: number }>(
      localPath,
      `SELECT watched FROM imm_videos WHERE video_key = 'showa-e1'`,
    );
    assert.equal(video?.watched, 1);

    // Both sessions now credit the same video; episode was started once and
    // completed once (by the remote session that watched it to the end).
    const global = queryOne<{
      episodes_started: number;
      episodes_completed: number;
      total_sessions: number;
    }>(
      localPath,
      'SELECT episodes_started, episodes_completed, total_sessions FROM imm_lifetime_global WHERE global_id = 1',
    );
    assert.equal(global?.total_sessions, 2);
    assert.equal(global?.episodes_started, 1);
    assert.equal(global?.episodes_completed, 1);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('backfills media metadata for matched videos when local metadata is absent', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(localPath, {
      uuid: 'local-1',
      videoKey: 'showa-e1',
      animeTitleKey: 'showa',
      startedAtMs: BASE_MS,
      applyLifetime: true,
    });
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'showa-e1',
      animeTitleKey: 'showa',
      startedAtMs: BASE_MS + DAY_MS,
      applyLifetime: true,
    });
    const remoteDb = new Database(remotePath, { readwrite: true });
    try {
      const remoteVideoId = Number(
        remoteDb
          .query<{
            video_id: number;
          }>(`SELECT video_id FROM imm_videos WHERE video_key = 'showa-e1'`)
          .get()?.video_id,
      );
      remoteDb
        .prepare(
          `INSERT INTO imm_media_art (video_id, anilist_id, cover_url, title_romaji, fetched_at_ms)
           VALUES (?, 123, 'https://example.test/cover.jpg', 'Show A', ?)`,
        )
        .run(remoteVideoId, String(BASE_MS));
      remoteDb
        .prepare(
          `INSERT INTO imm_youtube_videos (video_id, youtube_video_id, video_url, video_title, fetched_at_ms)
           VALUES (?, 'yt-1', 'https://youtube.test/watch?v=yt-1', 'Remote Video', ?)`,
        )
        .run(remoteVideoId, String(BASE_MS));
    } finally {
      remoteDb.close();
    }

    const summary = mergeSnapshotIntoDb(localPath, remotePath);
    assert.equal(summary.videosAdded, 0);
    const localVideoId = Number(
      queryOne<{ video_id: number }>(
        localPath,
        `SELECT video_id FROM imm_videos WHERE video_key = 'showa-e1'`,
      )?.video_id,
    );
    const art = queryOne<{ cover_url: string; title_romaji: string }>(
      localPath,
      'SELECT cover_url, title_romaji FROM imm_media_art WHERE video_id = ?',
      [localVideoId],
    );
    assert.equal(art?.cover_url, 'https://example.test/cover.jpg');
    assert.equal(art?.title_romaji, 'Show A');
    const youtube = queryOne<{ youtube_video_id: string; video_title: string }>(
      localPath,
      'SELECT youtube_video_id, video_title FROM imm_youtube_videos WHERE video_id = ?',
      [localVideoId],
    );
    assert.equal(youtube?.youtube_video_id, 'yt-1');
    assert.equal(youtube?.video_title, 'Remote Video');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('does not double-count active_days when a merged session starts earlier on an already-credited day', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(localPath, {
      uuid: 'local-1',
      videoKey: 'showa-e1',
      startedAtMs: BASE_MS + 3_600_000,
      applyLifetime: true,
    });
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'showb-e1',
      startedAtMs: BASE_MS,
      applyLifetime: true,
    });

    mergeSnapshotIntoDb(localPath, remotePath);
    const global = queryOne<{ active_days: number; total_sessions: number }>(
      localPath,
      'SELECT active_days, total_sessions FROM imm_lifetime_global WHERE global_id = 1',
    );
    assert.equal(global?.total_sessions, 2);
    assert.equal(global?.active_days, 1);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('skips unfinished sessions', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(remotePath, {
      uuid: 'remote-active',
      videoKey: 'showa-e1',
      startedAtMs: BASE_MS,
      endedAtMs: null,
    });

    const summary = mergeSnapshotIntoDb(localPath, remotePath);
    assert.equal(summary.sessionsMerged, 0);
    assert.equal(summary.activeSessionsSkipped, 1);
    assert.equal(count(localPath, 'SELECT COUNT(*) AS n FROM imm_sessions'), 0);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('copies remote-only historical rollups but never sums shared groups', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    // Remote has a video plus an old rollup row whose sessions were pruned.
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'old-show-e1',
      startedAtMs: BASE_MS,
      applyLifetime: true,
    });
    withWritableDb(remotePath, (remoteDb) => {
      const remoteVideoId = Number(
        remoteDb
          .query<{
            video_id: number;
          }>(`SELECT video_id FROM imm_videos WHERE video_key = 'old-show-e1'`)
          .get()?.video_id,
      );
      remoteDb
        .prepare(
          `INSERT INTO imm_daily_rollups (rollup_day, video_id, total_sessions, total_active_min, total_lines_seen, total_tokens_seen, total_cards)
           VALUES (10000, ?, 4, 120.5, 400, 3200, 9)`,
        )
        .run(remoteVideoId);
    });

    // Local already has its own rollup row for a shared group.
    insertFixtureSession(localPath, {
      uuid: 'local-1',
      videoKey: 'old-show-e1',
      startedAtMs: BASE_MS,
      applyLifetime: true,
    });
    const localVideoId = withWritableDb(localPath, (localDb) => {
      const videoId = Number(
        localDb
          .query<{
            video_id: number;
          }>(`SELECT video_id FROM imm_videos WHERE video_key = 'old-show-e1'`)
          .get()?.video_id,
      );
      localDb
        .prepare(
          `INSERT INTO imm_daily_rollups (rollup_day, video_id, total_sessions, total_active_min, total_lines_seen, total_tokens_seen, total_cards)
           VALUES (10000, ?, 2, 60.0, 200, 1600, 4)`,
        )
        .run(videoId);
      return videoId;
    });

    const summary = mergeSnapshotIntoDb(localPath, remotePath);
    // Shared group (day 10000) untouched; the remote-only session's own group
    // was recomputed, not copied.
    assert.equal(summary.dailyRollupsCopied, 0);
    const shared = queryOne<{ total_sessions: number }>(
      localPath,
      'SELECT total_sessions FROM imm_daily_rollups WHERE rollup_day = 10000 AND video_id = ?',
      [localVideoId],
    );
    assert.equal(shared?.total_sessions, 2);

    // Now a rollup for a video with no local sessions at all gets copied.
    withWritableDb(remotePath, (remoteDb2) => {
      const orphanVideoId = Number(
        remoteDb2
          .prepare(
            `INSERT INTO imm_videos (video_key, canonical_title, source_type, watched, duration_ms)
             VALUES ('pruned-show-e1', 'pruned-show-e1', 1, 1, 1440000)`,
          )
          .run().lastInsertRowid,
      );
      remoteDb2
        .prepare(
          `INSERT INTO imm_daily_rollups (rollup_day, video_id, total_sessions, total_active_min, total_lines_seen, total_tokens_seen, total_cards)
           VALUES (9000, ?, 3, 90.0, 300, 2400, 6)`,
        )
        .run(orphanVideoId);
    });

    const summary2 = mergeSnapshotIntoDb(localPath, remotePath);
    assert.equal(summary2.dailyRollupsCopied, 1);
    const copiedVideoId = Number(
      queryOne<{ video_id: number }>(
        localPath,
        `SELECT video_id FROM imm_videos WHERE video_key = 'pruned-show-e1'`,
      )?.video_id,
    );
    const copied = queryOne<{ total_sessions: number; total_active_min: number }>(
      localPath,
      'SELECT total_sessions, total_active_min FROM imm_daily_rollups WHERE rollup_day = 9000 AND video_id = ?',
      [copiedVideoId],
    );
    assert.equal(copied?.total_sessions, 3);
    assert.equal(copied?.total_active_min, 90.0);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('does not copy remote-only daily rollups into months with local sessions', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(localPath, {
      uuid: 'local-1',
      videoKey: 'mixed-month-e1',
      animeTitleKey: 'mixed-month',
      startedAtMs: BASE_MS,
      applyLifetime: true,
    });
    const remoteDb = new Database(remotePath, { readwrite: true });
    try {
      const remoteVideoId = Number(
        remoteDb
          .prepare(
            `INSERT INTO imm_videos (video_key, canonical_title, source_type, watched, duration_ms)
             VALUES ('mixed-month-e1', 'mixed-month-e1', 1, 1, 1440000)`,
          )
          .run().lastInsertRowid,
      );
      remoteDb
        .prepare(
          `INSERT INTO imm_daily_rollups (rollup_day, video_id, total_sessions, total_active_min, total_lines_seen, total_tokens_seen, total_cards)
           VALUES (20615, ?, 3, 90.0, 300, 2400, 6)`,
        )
        .run(remoteVideoId);
      remoteDb
        .prepare(
          `INSERT INTO imm_monthly_rollups (rollup_month, video_id, total_sessions, total_active_min, total_lines_seen, total_tokens_seen, total_cards)
           VALUES (202606, ?, 3, 90.0, 300, 2400, 6)`,
        )
        .run(remoteVideoId);
    } finally {
      remoteDb.close();
    }

    const summary = mergeSnapshotIntoDb(localPath, remotePath);
    assert.equal(summary.dailyRollupsCopied, 0);
    assert.equal(summary.monthlyRollupsCopied, 0);
    const localVideoId = Number(
      queryOne<{ video_id: number }>(
        localPath,
        `SELECT video_id FROM imm_videos WHERE video_key = 'mixed-month-e1'`,
      )?.video_id,
    );
    assert.equal(
      count(
        localPath,
        'SELECT COUNT(*) AS n FROM imm_daily_rollups WHERE rollup_day = 20615 AND video_id = ?',
        [localVideoId],
      ),
      0,
    );
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('rejects snapshots at a different schema version', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    withWritableDb(remotePath, (remoteDb) => {
      remoteDb.prepare('UPDATE imm_schema_version SET schema_version = 17').run();
    });

    assert.throws(() => mergeSnapshotIntoDb(localPath, remotePath), /schema version 17/);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('createDbSnapshot produces a mergeable copy', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'showa-e1',
      startedAtMs: BASE_MS,
      applyLifetime: true,
    });
    const snapshotPath = path.join(dir, 'snapshot.sqlite');
    createDbSnapshot(remotePath, snapshotPath);
    assert.ok(fs.existsSync(snapshotPath));

    const summary = mergeSnapshotIntoDb(localPath, snapshotPath);
    assert.equal(summary.sessionsMerged, 1);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('adopted word frequency excludes active-session counts that merge later', () => {
  const { dir, localPath, remotePath } = makeDbPair();
  try {
    // Remote has an ended session and a stale ACTIVE one (e.g. app crashed).
    // The remote tracker increments frequency live, so the word's frequency (5)
    // already includes the active session's 4 occurrences even though that
    // session's lines are skipped by the merge.
    insertFixtureSession(remotePath, {
      uuid: 'remote-ended',
      videoKey: 'showb-e1',
      animeTitleKey: 'showb',
      startedAtMs: BASE_MS,
      applyLifetime: true,
      words: [{ headword: '食べる', word: '食べた', reading: 'たべた', count: 1 }],
    });
    insertFixtureSession(remotePath, {
      uuid: 'remote-active',
      videoKey: 'showb-e2',
      animeTitleKey: 'showb',
      startedAtMs: BASE_MS + DAY_MS,
      endedAtMs: null,
      words: [{ headword: '食べる', word: '食べた', reading: 'たべた', count: 4 }],
    });

    const first = mergeSnapshotIntoDb(localPath, remotePath);
    assert.equal(first.sessionsMerged, 1);
    assert.equal(first.activeSessionsSkipped, 1);
    assert.equal(first.wordsAdded, 1);
    // Only the ended session's count is adopted; the active session's slice
    // is re-added when that session finalizes and syncs.
    assert.equal(
      queryOne<{ frequency: number }>(localPath, `SELECT frequency FROM imm_words WHERE word = '食べた'`)
        ?.frequency,
      1,
    );

    // The remote app restarts and finalizes the stale session.
    withWritableDb(remotePath, (db) => {
      db.prepare(
        `UPDATE imm_sessions SET ended_at_ms = ?, status = 2
         WHERE session_uuid = 'remote-active'`,
      ).run(String(BASE_MS + DAY_MS + 1_500_000));
    });

    const second = mergeSnapshotIntoDb(localPath, remotePath);
    assert.equal(second.sessionsMerged, 1);
    assert.equal(second.sessionsAlreadyPresent, 1);
    // 1 (ended session) + 4 (finalized session), not 5 + 4 = 9.
    assert.equal(
      queryOne<{ frequency: number }>(localPath, `SELECT frequency FROM imm_words WHERE word = '食べた'`)
        ?.frequency,
      5,
    );
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
