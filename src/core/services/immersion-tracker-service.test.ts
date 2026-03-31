import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { toMonthKey } from './immersion-tracker/maintenance';
import { enqueueWrite } from './immersion-tracker/queue';
import { toDbTimestamp } from './immersion-tracker/query-shared';
import { Database, type DatabaseSync } from './immersion-tracker/sqlite';
import { nowMs as trackerNowMs } from './immersion-tracker/time';
import {
  deriveCanonicalTitle,
  normalizeText,
  resolveBoundedInt,
} from './immersion-tracker/reducer';
import type { QueuedWrite } from './immersion-tracker/types';
import { PartOfSpeech, type MergedToken } from '../../types';

type ImmersionTrackerService = import('./immersion-tracker-service').ImmersionTrackerService;
type ImmersionTrackerServiceCtor =
  typeof import('./immersion-tracker-service').ImmersionTrackerService;

let trackerCtor: ImmersionTrackerServiceCtor | null = null;

async function loadTrackerCtor(): Promise<ImmersionTrackerServiceCtor> {
  if (trackerCtor) return trackerCtor;
  const mod = await import('./immersion-tracker-service');
  trackerCtor = mod.ImmersionTrackerService;
  return trackerCtor;
}

async function waitForPendingAnimeMetadata(tracker: ImmersionTrackerService): Promise<void> {
  const privateApi = tracker as unknown as {
    sessionState: { videoId: number } | null;
    pendingAnimeMetadataUpdates?: Map<number, Promise<void>>;
  };
  const videoId = privateApi.sessionState?.videoId;
  if (!videoId) return;
  await privateApi.pendingAnimeMetadataUpdates?.get(videoId);
}

async function waitForCondition(
  predicate: () => boolean,
  timeoutMs = 1_000,
  intervalMs = 10,
): Promise<void> {
  const start = globalThis.performance?.now() ?? 0;
  const deadline = start + timeoutMs;
  while ((globalThis.performance?.now() ?? deadline) < deadline) {
    if (predicate()) {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
  assert.equal(predicate(), true);
}

function makeMergedToken(overrides: Partial<MergedToken>): MergedToken {
  return {
    surface: '',
    reading: '',
    headword: '',
    startPos: 0,
    endPos: 0,
    partOfSpeech: PartOfSpeech.other,
    pos1: '',
    pos2: '',
    pos3: '',
    isMerged: true,
    isKnown: false,
    isNPlusOneTarget: false,
    ...overrides,
  };
}

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-immersion-test-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  const dir = path.dirname(dbPath);
  if (!fs.existsSync(dir)) {
    return;
  }

  const bunRuntime = globalThis as typeof globalThis & {
    Bun?: {
      gc?: (force?: boolean) => void;
    };
  };
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      fs.rmSync(dir, { recursive: true, force: true });
      return;
    } catch (error) {
      const err = error as NodeJS.ErrnoException;
      if (process.platform !== 'win32' || err.code !== 'EBUSY') {
        throw error;
      }
      bunRuntime.Bun?.gc?.(true);
      Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 25);
    }
  }

  // libsql keeps Windows file handles alive after close when prepared statements were used.
}

test('seam: resolveBoundedInt keeps fallback for invalid values', () => {
  assert.equal(resolveBoundedInt(undefined, 25, 1, 100), 25);
  assert.equal(resolveBoundedInt(0, 25, 1, 100), 25);
  assert.equal(resolveBoundedInt(101, 25, 1, 100), 25);
  assert.equal(resolveBoundedInt(44.8, 25, 1, 100), 44);
});

test('seam: reducer title normalization covers local and remote paths', () => {
  assert.equal(normalizeText('  hello\n  world  '), 'hello world');
  assert.equal(deriveCanonicalTitle('/tmp/Episode 01.mkv'), 'Episode 01');
  assert.equal(
    deriveCanonicalTitle('https://cdn.example.com/show/%E7%AC%AC1%E8%A9%B1.mp4'),
    '\u7b2c1\u8a71',
  );
});

test('seam: enqueueWrite drops oldest entries once capacity is exceeded', () => {
  const queue: QueuedWrite[] = [
    { kind: 'event', sessionId: 1, eventType: 1, sampleMs: 1000 },
    { kind: 'event', sessionId: 1, eventType: 2, sampleMs: 1001 },
  ];
  const incoming: QueuedWrite = { kind: 'event', sessionId: 1, eventType: 3, sampleMs: 1002 };

  const result = enqueueWrite(queue, incoming, 2);
  assert.equal(result.dropped, 1);
  assert.equal(queue.length, 2);
  assert.equal((queue[0] as Extract<QueuedWrite, { kind: 'event' }>).eventType, 2);
  assert.equal((queue[1] as Extract<QueuedWrite, { kind: 'event' }>).eventType, 3);
});

test('seam: toMonthKey uses UTC calendar month', () => {
  assert.equal(toMonthKey(-86_400_000), 196912);
  assert.equal(toMonthKey(0), 197001);
});

test('startSession generates UUID-like session identifiers', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/episode.mkv', 'Episode');

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const db = new Database(dbPath);
    const row = db.prepare('SELECT session_uuid FROM imm_sessions LIMIT 1').get() as {
      session_uuid: string;
    } | null;
    db.close();

    assert.equal(typeof row?.session_uuid, 'string');
    assert.equal(row?.session_uuid?.startsWith('session-'), false);
    assert.ok(/^[0-9a-fA-F-]{36}$/.test(row?.session_uuid || ''));
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('destroy finalizes active session and persists final telemetry', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/episode-2.mkv', 'Episode 2');
    tracker.recordSubtitleLine('Hello immersion', 0, 1);
    tracker.destroy();

    const db = new Database(dbPath);
    const sessionRow = db.prepare('SELECT ended_at_ms FROM imm_sessions LIMIT 1').get() as {
      ended_at_ms: string | number | null;
    } | null;
    const telemetryCountRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_session_telemetry')
      .get() as { total: number };
    db.close();

    assert.ok(sessionRow);
    assert.notEqual(sessionRow?.ended_at_ms, null);
    assert.ok(Number(telemetryCountRow.total) >= 2);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('finalize updates lifetime summary rows from final session metrics', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/Little Witch Academia S02E05.mkv', 'Episode 5');
    await waitForPendingAnimeMetadata(tracker);

    const privateApi = tracker as unknown as {
      sessionState: { sessionId: number; videoId: number } | null;
    };
    const sessionId = privateApi.sessionState?.sessionId;
    const videoId = privateApi.sessionState?.videoId;
    assert.ok(sessionId);
    assert.ok(videoId);

    tracker.recordCardsMined(2);
    tracker.recordSubtitleLine('today is bright', 0, 1.2);
    tracker.recordLookup(true);

    tracker.destroy();

    const db = new Database(dbPath);
    const globalRow = db
      .prepare('SELECT total_sessions, total_cards, total_active_ms FROM imm_lifetime_global')
      .get() as {
      total_sessions: number;
      total_cards: number;
      total_active_ms: number;
    } | null;
    const mediaRow = db
      .prepare(
        'SELECT total_sessions, total_cards, total_active_ms, total_tokens_seen, total_lines_seen FROM imm_lifetime_media WHERE video_id = ?',
      )
      .get(videoId) as {
      total_sessions: number;
      total_cards: number;
      total_active_ms: number;
      total_tokens_seen: number;
      total_lines_seen: number;
    } | null;
    const animeIdRow = db
      .prepare('SELECT anime_id FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { anime_id: number | null } | null;
    const animeRow = animeIdRow?.anime_id
      ? (db
          .prepare('SELECT total_sessions, total_cards FROM imm_lifetime_anime WHERE anime_id = ?')
          .get(animeIdRow.anime_id) as {
          total_sessions: number;
          total_cards: number;
        } | null)
      : null;
    const appliedRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_lifetime_applied_sessions WHERE session_id = ?')
      .get(sessionId) as {
      total: number;
    } | null;
    db.close();

    assert.ok(globalRow);
    assert.equal(globalRow?.total_sessions, 1);
    assert.equal(globalRow?.total_cards, 2);
    assert.ok(Number(globalRow?.total_active_ms ?? 0) >= 0);
    assert.ok(mediaRow);
    assert.equal(mediaRow?.total_sessions, 1);
    assert.equal(mediaRow?.total_cards, 2);
    assert.equal(mediaRow?.total_lines_seen, 1);
    assert.ok(animeRow);
    assert.equal(animeRow?.total_sessions, 1);
    assert.equal(animeRow?.total_cards, 2);
    assert.equal(appliedRow?.total, 1);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('lifetime updates are not double-counted if finalize runs multiple times', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/Little Witch Academia S02E06.mkv', 'Episode 6');
    await waitForPendingAnimeMetadata(tracker);

    const privateApi = tracker as unknown as {
      finalizeActiveSession: () => void;
      sessionState: { sessionId: number; videoId: number } | null;
    };
    const sessionState = privateApi.sessionState;
    const sessionId = sessionState?.sessionId;
    assert.ok(sessionId);

    tracker.recordCardsMined(3);
    privateApi.finalizeActiveSession();
    privateApi.sessionState = sessionState;
    privateApi.finalizeActiveSession();

    const db = new Database(dbPath);
    const globalRow = db
      .prepare('SELECT total_sessions, total_cards FROM imm_lifetime_global')
      .get() as {
      total_sessions: number;
      total_cards: number;
    } | null;
    const appliedRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_lifetime_applied_sessions WHERE session_id = ?')
      .get(sessionId) as {
      total: number;
    } | null;
    db.close();

    assert.ok(globalRow);
    assert.equal(globalRow?.total_sessions, 1);
    assert.equal(globalRow?.total_cards, 3);
    assert.equal(appliedRow?.total, 1);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('lifetime counters use distinct-day and distinct-video semantics', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/Little Witch Academia S02E05.mkv', 'Episode 5');
    await waitForPendingAnimeMetadata(tracker);
    let privateApi = tracker as unknown as {
      db: DatabaseSync;
      sessionState: { sessionId: number; videoId: number } | null;
    };
    const firstVideoId = privateApi.sessionState?.videoId;
    assert.ok(firstVideoId);
    const animeId = (
      privateApi.db
        .prepare('SELECT anime_id FROM imm_videos WHERE video_id = ?')
        .get(firstVideoId) as {
        anime_id: number | null;
      } | null
    )?.anime_id;
    assert.ok(animeId);
    privateApi.db
      .prepare('UPDATE imm_anime SET episodes_total = 2 WHERE anime_id = ?')
      .run(animeId);
    await tracker.setVideoWatched(firstVideoId, true);
    tracker.destroy();

    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/Little Witch Academia S02E05.mkv', 'Episode 5');
    await waitForPendingAnimeMetadata(tracker);
    privateApi = tracker as unknown as typeof privateApi;
    const repeatedSessionApi = tracker as unknown as {
      sessionState: { sessionId: number; videoId: number } | null;
    };
    const repeatedVideoId = repeatedSessionApi.sessionState?.videoId;
    assert.equal(repeatedVideoId, firstVideoId);
    await tracker.setVideoWatched(repeatedVideoId, true);
    tracker.destroy();

    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/Little Witch Academia S02E06.mkv', 'Episode 6');
    await waitForPendingAnimeMetadata(tracker);
    privateApi = tracker as unknown as typeof privateApi;
    const secondSessionApi = tracker as unknown as {
      sessionState: { sessionId: number; videoId: number } | null;
    };
    const secondVideoId = secondSessionApi.sessionState?.videoId;
    assert.ok(secondVideoId);
    assert.ok(secondVideoId !== firstVideoId);
    await tracker.setVideoWatched(secondVideoId, true);
    tracker.destroy();

    const db = new Database(dbPath);
    const globalRow = db
      .prepare(
        'SELECT total_sessions, active_days, episodes_started, episodes_completed, anime_completed FROM imm_lifetime_global',
      )
      .get() as {
      total_sessions: number;
      active_days: number;
      episodes_started: number;
      episodes_completed: number;
      anime_completed: number;
    } | null;
    const firstMediaRow = db
      .prepare('SELECT completed FROM imm_lifetime_media WHERE video_id = ?')
      .get(firstVideoId) as { completed: number } | null;
    const secondMediaRow = db
      .prepare('SELECT completed FROM imm_lifetime_media WHERE video_id = ?')
      .get(secondVideoId) as { completed: number } | null;
    const animeRow = db
      .prepare(
        'SELECT episodes_started, episodes_completed FROM imm_lifetime_anime WHERE anime_id = ?',
      )
      .get(animeId) as { episodes_started: number; episodes_completed: number } | null;
    db.close();

    assert.ok(globalRow);
    assert.equal(globalRow?.total_sessions, 3);
    assert.equal(globalRow?.active_days, 1);
    assert.equal(globalRow?.episodes_started, 2);
    assert.equal(globalRow?.episodes_completed, 2);
    assert.equal(globalRow?.anime_completed, 1);
    assert.ok(firstMediaRow);
    assert.equal(firstMediaRow?.completed, 1);
    assert.ok(secondMediaRow);
    assert.equal(secondMediaRow?.completed, 1);
    assert.ok(animeRow);
    assert.equal(animeRow?.episodes_started, 2);
    assert.equal(animeRow?.episodes_completed, 2);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('rebuildLifetimeSummaries backfills retained ended sessions and resets stale lifetime rows', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/Little Witch Academia S02E05.mkv', 'Episode 5');
    await waitForPendingAnimeMetadata(tracker);
    const firstApi = tracker as unknown as {
      db: DatabaseSync;
      sessionState: { videoId: number } | null;
    };
    const firstVideoId = firstApi.sessionState?.videoId;
    if (firstVideoId == null) {
      throw new Error('Expected first session video id');
    }
    const animeId = (
      firstApi.db
        .prepare('SELECT anime_id FROM imm_videos WHERE video_id = ?')
        .get(firstVideoId) as {
        anime_id: number | null;
      } | null
    )?.anime_id;
    assert.ok(animeId);
    firstApi.db.prepare('UPDATE imm_anime SET episodes_total = 2 WHERE anime_id = ?').run(animeId);
    tracker.recordCardsMined(2);
    await tracker.setVideoWatched(firstVideoId, true);
    tracker.destroy();

    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/Little Witch Academia S02E06.mkv', 'Episode 6');
    await waitForPendingAnimeMetadata(tracker);
    const secondApi = tracker as unknown as {
      sessionState: { videoId: number } | null;
    };
    const secondVideoId = secondApi.sessionState?.videoId;
    if (secondVideoId == null) {
      throw new Error('Expected second session video id');
    }
    tracker.recordCardsMined(1);
    await tracker.setVideoWatched(secondVideoId, true);
    tracker.destroy();

    tracker = new Ctor({ dbPath });
    const rebuildApi = tracker as unknown as { db: DatabaseSync };
    rebuildApi.db
      .prepare(
        `
      UPDATE imm_lifetime_global
      SET
        total_sessions = 99,
        total_cards = 77,
        episodes_started = 88,
        episodes_completed = 66
      WHERE global_id = 1
      `,
      )
      .run();
    rebuildApi.db.exec(`
      DELETE FROM imm_lifetime_media;
      DELETE FROM imm_lifetime_anime;
      DELETE FROM imm_lifetime_applied_sessions;
    `);

    const rebuild = await tracker.rebuildLifetimeSummaries();

    const globalRow = rebuildApi.db
      .prepare(
        'SELECT total_sessions, total_cards, episodes_started, episodes_completed, anime_completed, last_rebuilt_ms FROM imm_lifetime_global WHERE global_id = 1',
      )
      .get() as {
      total_sessions: number;
      total_cards: number;
      episodes_started: number;
      episodes_completed: number;
      anime_completed: number;
      last_rebuilt_ms: string | number | null;
    } | null;
    const appliedSessions = rebuildApi.db
      .prepare('SELECT COUNT(*) AS total FROM imm_lifetime_applied_sessions')
      .get() as { total: number } | null;

    assert.equal(rebuild.appliedSessions, 2);
    assert.ok(rebuild.rebuiltAtMs > 0);
    assert.ok(globalRow);
    assert.equal(globalRow?.total_sessions, 2);
    assert.equal(globalRow?.total_cards, 3);
    assert.equal(globalRow?.episodes_started, 2);
    assert.equal(globalRow?.episodes_completed, 2);
    assert.equal(globalRow?.anime_completed, 1);
    assert.equal(globalRow?.last_rebuilt_ms, toDbTimestamp(rebuild.rebuiltAtMs));
    assert.equal(appliedSessions?.total, 2);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('fresh tracker DB creates lifetime summary tables', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    const db = new Database(dbPath);
    const tableRows = db
      .prepare("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
      .all() as Array<{ name: string }>;
    db.close();

    const tableNames = new Set(tableRows.map((row) => row.name));
    const expectedTables = [
      'imm_lifetime_global',
      'imm_lifetime_anime',
      'imm_lifetime_media',
      'imm_lifetime_applied_sessions',
    ];

    for (const tableName of expectedTables) {
      assert.ok(tableNames.has(tableName), `Expected ${tableName} to exist`);
    }
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('startup backfills lifetime summaries when retained sessions exist but summary tables are empty', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/KonoSuba S02E05.mkv', 'Episode 5');
    await waitForPendingAnimeMetadata(tracker);
    tracker.recordCardsMined(2);
    tracker.destroy();

    const db = new Database(dbPath);
    db.exec(`
      DELETE FROM imm_lifetime_media;
      DELETE FROM imm_lifetime_anime;
      DELETE FROM imm_lifetime_applied_sessions;
      UPDATE imm_lifetime_global
      SET
        total_sessions = 0,
        total_active_ms = 0,
        total_cards = 0,
        active_days = 0,
        episodes_started = 0,
        episodes_completed = 0,
        anime_completed = 0
      WHERE global_id = 1;
    `);
    db.close();

    tracker = new Ctor({ dbPath });
    const trackerApi = tracker as unknown as { db: DatabaseSync };
    const globalRow = trackerApi.db
      .prepare(
        'SELECT total_sessions, total_cards, active_days FROM imm_lifetime_global WHERE global_id = 1',
      )
      .get() as {
      total_sessions: number;
      total_cards: number;
      active_days: number;
    } | null;
    const mediaRows = trackerApi.db
      .prepare('SELECT COUNT(*) AS total FROM imm_lifetime_media')
      .get() as { total: number } | null;
    const appliedRows = trackerApi.db
      .prepare('SELECT COUNT(*) AS total FROM imm_lifetime_applied_sessions')
      .get() as { total: number } | null;

    assert.ok(globalRow);
    assert.equal(globalRow?.total_sessions, 1);
    assert.equal(globalRow?.total_cards, 2);
    assert.equal(globalRow?.active_days, 1);
    assert.equal(mediaRows?.total, 1);
    assert.equal(appliedRows?.total, 1);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('startup finalizes stale active sessions and applies lifetime summaries', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const trackerApi = tracker as unknown as { db: DatabaseSync };
    const db = trackerApi.db;
    const startedAtMs = trackerNowMs() - 10_000;
    const sampleMs = startedAtMs + 5_000;

    db.prepare(
      `
        INSERT INTO imm_anime (
          anime_id,
          canonical_title,
          normalized_title_key,
          episodes_total,
          CREATED_DATE,
          LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?)
      `,
    ).run(1, 'KonoSuba', 'konosuba', 10, toDbTimestamp(startedAtMs), toDbTimestamp(startedAtMs));

    db.prepare(
      `
        INSERT INTO imm_videos (
          video_id,
          video_key,
          canonical_title,
          anime_id,
          watched,
          source_type,
          duration_ms,
          CREATED_DATE,
          LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(
      1,
      'local:/tmp/konosuba-s02e05.mkv',
      'KonoSuba S02E05',
      1,
      1,
      1,
      0,
      toDbTimestamp(startedAtMs),
      toDbTimestamp(startedAtMs),
    );

    db.prepare(
      `
        INSERT INTO imm_sessions (
          session_id,
          session_uuid,
          video_id,
          started_at_ms,
          status,
          ended_media_ms,
          CREATED_DATE,
          LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(
      1,
      '11111111-1111-1111-1111-111111111111',
      1,
      toDbTimestamp(startedAtMs),
      1,
      321000,
      toDbTimestamp(startedAtMs),
      toDbTimestamp(sampleMs),
    );

    db.prepare(
      `
        INSERT INTO imm_session_telemetry (
          session_id,
          sample_ms,
          total_watched_ms,
          active_watched_ms,
          lines_seen,
          tokens_seen,
          cards_mined,
          lookup_count,
          lookup_hits,
          pause_count,
          pause_ms,
          seek_forward_count,
          seek_backward_count,
          media_buffer_events
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(1, toDbTimestamp(sampleMs), 5000, 4000, 12, 120, 2, 5, 3, 1, 250, 1, 0, 0);

    tracker.destroy();
    tracker = new Ctor({ dbPath });

    const restartedApi = tracker as unknown as { db: DatabaseSync };
    const sessionRow = restartedApi.db
      .prepare(
        `
          SELECT ended_at_ms, status, ended_media_ms, active_watched_ms, tokens_seen, cards_mined
          FROM imm_sessions
          WHERE session_id = 1
        `,
      )
      .get() as {
      ended_at_ms: string | number | null;
      status: number;
      ended_media_ms: number | null;
      active_watched_ms: number;
      tokens_seen: number;
      cards_mined: number;
    } | null;
    const globalRow = restartedApi.db
      .prepare(
        `
          SELECT total_sessions, total_active_ms, total_cards, active_days, episodes_started,
                 episodes_completed
          FROM imm_lifetime_global
          WHERE global_id = 1
        `,
      )
      .get() as {
      total_sessions: number;
      total_active_ms: number;
      total_cards: number;
      active_days: number;
      episodes_started: number;
      episodes_completed: number;
    } | null;
    const mediaRows = restartedApi.db
      .prepare('SELECT COUNT(*) AS total FROM imm_lifetime_media')
      .get() as { total: number } | null;
    const animeRows = restartedApi.db
      .prepare('SELECT COUNT(*) AS total FROM imm_lifetime_anime')
      .get() as { total: number } | null;
    const appliedRows = restartedApi.db
      .prepare('SELECT COUNT(*) AS total FROM imm_lifetime_applied_sessions')
      .get() as { total: number } | null;

    assert.ok(sessionRow);
    assert.equal(sessionRow?.ended_at_ms, toDbTimestamp(sampleMs));
    assert.equal(sessionRow?.status, 2);
    assert.equal(sessionRow?.ended_media_ms, 321_000);
    assert.equal(sessionRow?.active_watched_ms, 4000);
    assert.equal(sessionRow?.tokens_seen, 120);
    assert.equal(sessionRow?.cards_mined, 2);

    assert.ok(globalRow);
    assert.equal(globalRow?.total_sessions, 1);
    assert.equal(globalRow?.total_active_ms, 4000);
    assert.equal(globalRow?.total_cards, 2);
    assert.equal(globalRow?.active_days, 1);
    assert.equal(globalRow?.episodes_started, 1);
    assert.equal(globalRow?.episodes_completed, 1);
    assert.equal(mediaRows?.total, 1);
    assert.equal(animeRows?.total, 1);
    assert.equal(appliedRows?.total, 1);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('persists and retrieves minimum immersion tracking fields', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/episode-3.mkv', 'Episode 3');
    tracker.recordSubtitleLine('alpha beta', 0, 1.2, [
      makeMergedToken({
        surface: 'alpha',
        headword: 'alpha',
        reading: 'alpha',
      }),
      makeMergedToken({
        surface: 'beta',
        headword: 'beta',
        reading: 'beta',
      }),
    ]);
    tracker.recordCardsMined(2);
    tracker.recordLookup(true);
    tracker.recordPlaybackPosition(12.5);

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const summaries = await tracker.getSessionSummaries(10);
    assert.ok(summaries.length >= 1);
    assert.ok(summaries[0]!.linesSeen >= 1);
    assert.ok(summaries[0]!.cardsMined >= 2);

    tracker.destroy();

    const db = new Database(dbPath);
    const videoRow = db
      .prepare('SELECT canonical_title, source_path, duration_ms FROM imm_videos LIMIT 1')
      .get() as {
      canonical_title: string;
      source_path: string | null;
      duration_ms: number;
    } | null;
    const telemetryRow = db
      .prepare(
        `SELECT lines_seen, tokens_seen, cards_mined
         FROM imm_session_telemetry
         ORDER BY sample_ms DESC, telemetry_id DESC
         LIMIT 1`,
      )
      .get() as {
      lines_seen: number;
      tokens_seen: number;
      cards_mined: number;
    } | null;
    db.close();

    assert.ok(videoRow);
    assert.equal(videoRow?.canonical_title, 'Episode 3');
    assert.equal(videoRow?.source_path, '/tmp/episode-3.mkv');
    assert.ok(Number(videoRow?.duration_ms ?? -1) >= 0);

    assert.ok(telemetryRow);
    assert.ok(Number(telemetryRow?.lines_seen ?? 0) >= 1);
    assert.ok(Number(telemetryRow?.tokens_seen ?? 0) >= 2);
    assert.ok(Number(telemetryRow?.cards_mined ?? 0) >= 2);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('recordYomitanLookup persists a dedicated lookup counter without changing annotation lookup metrics', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/episode-yomitan.mkv', 'Episode Yomitan');
    tracker.recordSubtitleLine('alpha beta gamma', 0, 1.2);
    tracker.recordLookup(true);
    tracker.recordYomitanLookup();

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const summaries = await tracker.getSessionSummaries(10);
    assert.ok(summaries.length >= 1);
    assert.equal(summaries[0]?.lookupCount, 1);
    assert.equal(summaries[0]?.lookupHits, 1);
    assert.equal(summaries[0]?.yomitanLookupCount, 1);

    tracker.destroy();

    const db = new Database(dbPath);
    const sessionRow = db
      .prepare('SELECT lookup_count, lookup_hits, yomitan_lookup_count FROM imm_sessions LIMIT 1')
      .get() as {
      lookup_count: number;
      lookup_hits: number;
      yomitan_lookup_count: number;
    } | null;
    const eventRow = db
      .prepare(
        'SELECT event_type FROM imm_session_events WHERE event_type = ? ORDER BY ts_ms DESC LIMIT 1',
      )
      .get(9) as { event_type: number } | null;
    db.close();

    assert.equal(sessionRow?.lookup_count, 1);
    assert.equal(sessionRow?.lookup_hits, 1);
    assert.equal(sessionRow?.yomitan_lookup_count, 1);
    assert.equal(eventRow?.event_type, 9);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('recordSubtitleLine persists counted allowed tokenized vocabulary rows and subtitle-line occurrences', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/Little Witch Academia S02E04.mkv', 'Episode 4');
    await waitForPendingAnimeMetadata(tracker);
    tracker.recordSubtitleLine('猫 猫 日 日 は 知っている', 0, 1, [
      makeMergedToken({
        surface: '猫',
        headword: '猫',
        reading: 'ねこ',
        partOfSpeech: PartOfSpeech.noun,
        pos1: '名詞',
        pos2: '一般',
      }),
      makeMergedToken({
        surface: '猫',
        headword: '猫',
        reading: 'ねこ',
        partOfSpeech: PartOfSpeech.noun,
        pos1: '名詞',
        pos2: '一般',
      }),
      makeMergedToken({
        surface: 'は',
        headword: 'は',
        reading: 'は',
        partOfSpeech: PartOfSpeech.particle,
        pos1: '助詞',
        pos2: '係助詞',
      }),
      makeMergedToken({
        surface: '知っている',
        headword: '知る',
        reading: 'しっている',
        partOfSpeech: PartOfSpeech.other,
        pos1: '動詞',
        pos2: '自立',
      }),
    ]);

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const db = new Database(dbPath);
    const rows = db
      .prepare(
        `SELECT headword, word, reading, part_of_speech, pos1, pos2, frequency
         FROM imm_words
         ORDER BY id ASC`,
      )
      .all() as Array<{
      headword: string;
      word: string;
      reading: string;
      part_of_speech: string;
      pos1: string;
      pos2: string;
      frequency: number;
    }>;
    const lineRows = db
      .prepare(
        `SELECT video_id, anime_id, line_index, segment_start_ms, segment_end_ms, text
         FROM imm_subtitle_lines
         ORDER BY line_id ASC`,
      )
      .all() as Array<{
      video_id: number;
      anime_id: number | null;
      line_index: number;
      segment_start_ms: number | null;
      segment_end_ms: number | null;
      text: string;
    }>;
    const wordOccurrenceRows = db
      .prepare(
        `SELECT o.occurrence_count, w.headword, w.word, w.reading
         FROM imm_word_line_occurrences o
         JOIN imm_words w ON w.id = o.word_id
         ORDER BY o.line_id ASC, o.word_id ASC`,
      )
      .all() as Array<{
      occurrence_count: number;
      headword: string;
      word: string;
      reading: string;
    }>;
    const kanjiOccurrenceRows = db
      .prepare(
        `SELECT o.occurrence_count, k.kanji
         FROM imm_kanji_line_occurrences o
         JOIN imm_kanji k ON k.id = o.kanji_id
         ORDER BY o.line_id ASC, k.kanji ASC`,
      )
      .all() as Array<{
      occurrence_count: number;
      kanji: string;
    }>;
    db.close();

    assert.deepEqual(rows, [
      {
        headword: '猫',
        word: '猫',
        reading: 'ねこ',
        part_of_speech: PartOfSpeech.noun,
        pos1: '名詞',
        pos2: '一般',
        frequency: 2,
      },
      {
        headword: '知る',
        word: '知っている',
        reading: 'しっている',
        part_of_speech: PartOfSpeech.verb,
        pos1: '動詞',
        pos2: '自立',
        frequency: 1,
      },
    ]);
    assert.equal(lineRows.length, 1);
    assert.equal(lineRows[0]?.line_index, 1);
    assert.equal(lineRows[0]?.segment_start_ms, 0);
    assert.equal(lineRows[0]?.segment_end_ms, 1000);
    assert.equal(lineRows[0]?.text, '猫 猫 日 日 は 知っている');
    assert.ok(lineRows[0]?.video_id);
    assert.ok(lineRows[0]?.anime_id);
    assert.deepEqual(wordOccurrenceRows, [
      {
        occurrence_count: 2,
        headword: '猫',
        word: '猫',
        reading: 'ねこ',
      },
      {
        occurrence_count: 1,
        headword: '知る',
        word: '知っている',
        reading: 'しっている',
      },
    ]);
    assert.deepEqual(kanjiOccurrenceRows, [
      {
        occurrence_count: 2,
        kanji: '日',
      },
      {
        occurrence_count: 2,
        kanji: '猫',
      },
      {
        occurrence_count: 1,
        kanji: '知',
      },
    ]);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('recordSubtitleLine counts exact Yomitan tokens for session metrics', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/token-counting.mkv', 'Token Counting');
    tracker.recordSubtitleLine('猫 猫 日 日 は 知っている', 0, 1, [
      makeMergedToken({
        surface: '猫',
        headword: '猫',
        reading: 'ねこ',
        partOfSpeech: PartOfSpeech.noun,
        pos1: '名詞',
      }),
      makeMergedToken({
        surface: '猫',
        headword: '猫',
        reading: 'ねこ',
        partOfSpeech: PartOfSpeech.noun,
        pos1: '名詞',
      }),
      makeMergedToken({
        surface: 'は',
        headword: 'は',
        reading: 'は',
        partOfSpeech: PartOfSpeech.particle,
        pos1: '助詞',
      }),
      makeMergedToken({
        surface: '知っている',
        headword: '知る',
        reading: 'しっている',
        partOfSpeech: PartOfSpeech.other,
        pos1: '動詞',
      }),
    ]);

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const summaries = await tracker.getSessionSummaries(10);
    assert.equal(summaries[0]?.tokensSeen, 4);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('recordSubtitleLine leaves session token counts at zero when tokenization is unavailable', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/no-tokenization.mkv', 'No Tokenization');
    tracker.recordSubtitleLine('alpha beta gamma', 0, 1.2, null);

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const summaries = await tracker.getSessionSummaries(10);
    assert.equal(summaries[0]?.tokensSeen, 0);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('subtitle-line event payload omits duplicated subtitle text', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/payload-dup-test.mkv', 'Payload Dup Test');
    tracker.recordSubtitleLine('same line text', 0, 1);

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
      db: DatabaseSync;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const row = privateApi.db
      .prepare(
        `
          SELECT payload_json AS payloadJson
          FROM imm_session_events
          WHERE event_type = ?
          ORDER BY event_id DESC
          LIMIT 1
        `,
      )
      .get(1) as { payloadJson: string | null } | null;
    assert.ok(row?.payloadJson);
    const parsed = JSON.parse(row?.payloadJson ?? '{}') as {
      event?: string;
      tokens?: number;
      text?: string;
    };
    assert.equal(parsed.event, 'subtitle-line');
    assert.equal(typeof parsed.tokens, 'number');
    assert.equal('text' in parsed, false);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('recordPlaybackPosition marks watched at 85% completion', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/episode-85.mkv', 'Episode 85');
    tracker.recordMediaDuration(100);
    await waitForPendingAnimeMetadata(tracker);

    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      sessionState: { videoId: number } | null;
    };
    const videoId = privateApi.sessionState?.videoId;
    assert.ok(videoId);

    tracker.recordPlaybackPosition(84);
    let row = privateApi.db
      .prepare('SELECT watched FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { watched: number } | null;
    assert.equal(row?.watched, 0);

    tracker.recordPlaybackPosition(85);
    row = privateApi.db
      .prepare('SELECT watched FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { watched: number } | null;
    assert.equal(row?.watched, 1);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('flushTelemetry checkpoints latest playback position on the active session row', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange(
      '/tmp/episode-progress-checkpoint.mkv',
      'Episode Progress Checkpoint',
    );
    tracker.recordPlaybackPosition(91);

    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      sessionState: { sessionId: number } | null;
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    const sessionId = privateApi.sessionState?.sessionId;
    assert.ok(sessionId);

    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const row = privateApi.db
      .prepare('SELECT ended_media_ms FROM imm_sessions WHERE session_id = ?')
      .get(sessionId) as { ended_media_ms: number | null } | null;

    assert.ok(row);
    assert.equal(row?.ended_media_ms, 91_000);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('recordSubtitleLine advances session checkpoint progress when playback position is unavailable', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange(
      'https://stream.example.com/subtitle-progress.m3u8',
      'Subtitle Progress',
    );
    tracker.recordSubtitleLine('line one', 170, 185, [], null);

    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      sessionState: { sessionId: number } | null;
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    const sessionId = privateApi.sessionState?.sessionId;
    assert.ok(sessionId);

    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const row = privateApi.db
      .prepare('SELECT ended_media_ms FROM imm_sessions WHERE session_id = ?')
      .get(sessionId) as { ended_media_ms: number | null } | null;

    assert.equal(row?.ended_media_ms, 185_000);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('deleteSession ignores the currently active session and keeps new writes flushable', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/active-delete-test.mkv', 'Active Delete Test');

    const privateApi = tracker as unknown as {
      sessionState: { sessionId: number } | null;
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
      queue: unknown[];
    };
    const sessionId = privateApi.sessionState?.sessionId;
    assert.ok(sessionId);

    tracker.recordSubtitleLine('before delete', 0, 1);
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    await tracker.deleteSession(sessionId);

    tracker.recordSubtitleLine('after delete', 1, 2);
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const db = new Database(dbPath);
    const sessionCountRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_sessions WHERE session_id = ?')
      .get(sessionId) as { total: number };
    const subtitleLineCountRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_subtitle_lines WHERE session_id = ?')
      .get(sessionId) as { total: number };
    db.close();

    assert.equal(sessionCountRow.total, 1);
    assert.equal(subtitleLineCountRow.total, 2);
    assert.equal(privateApi.queue.length, 0);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('deleteVideo ignores the currently active video and keeps new writes flushable', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/active-video-delete-test.mkv', 'Active Video Delete Test');

    const privateApi = tracker as unknown as {
      sessionState: { sessionId: number; videoId: number } | null;
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
      queue: unknown[];
    };
    const sessionId = privateApi.sessionState?.sessionId;
    const videoId = privateApi.sessionState?.videoId;
    assert.ok(sessionId);
    assert.ok(videoId);

    tracker.recordSubtitleLine('before video delete', 0, 1);
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    await tracker.deleteVideo(videoId);

    tracker.recordSubtitleLine('after video delete', 1, 2);
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const db = new Database(dbPath);
    const sessionCountRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_sessions WHERE session_id = ?')
      .get(sessionId) as { total: number };
    const videoCountRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { total: number };
    const subtitleLineCountRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_subtitle_lines WHERE session_id = ?')
      .get(sessionId) as { total: number };
    db.close();

    assert.equal(sessionCountRow.total, 1);
    assert.equal(videoCountRow.total, 1);
    assert.equal(subtitleLineCountRow.total, 2);
    assert.equal(privateApi.queue.length, 0);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('handleMediaChange links parsed anime metadata on the active video row', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/Little Witch Academia S02E05.mkv', 'Episode 5');
    await waitForPendingAnimeMetadata(tracker);

    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      sessionState: { videoId: number } | null;
    };
    const videoId = privateApi.sessionState?.videoId;
    assert.ok(videoId);

    const row = privateApi.db
      .prepare(
        `
          SELECT
            v.anime_id,
            v.parsed_basename,
            v.parsed_title,
            v.parsed_season,
            v.parsed_episode,
            v.parser_source,
            a.canonical_title AS anime_title,
            a.anilist_id
          FROM imm_videos v
          LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
          WHERE v.video_id = ?
        `,
      )
      .get(videoId) as {
      anime_id: number | null;
      parsed_basename: string | null;
      parsed_title: string | null;
      parsed_season: number | null;
      parsed_episode: number | null;
      parser_source: string | null;
      anime_title: string | null;
      anilist_id: number | null;
    } | null;

    assert.ok(row);
    assert.ok(row?.anime_id);
    assert.equal(row?.parsed_basename, 'Little Witch Academia S02E05.mkv');
    assert.equal(row?.parsed_title, 'Little Witch Academia');
    assert.equal(row?.parsed_season, 2);
    assert.equal(row?.parsed_episode, 5);
    assert.ok(row?.parser_source === 'guessit' || row?.parser_source === 'fallback');
    assert.equal(row?.anime_title, 'Little Witch Academia');
    assert.equal(row?.anilist_id, null);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('handleMediaChange reuses the same provisional anime row across matching files', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/Little Witch Academia S02E05.mkv', 'Episode 5');
    await waitForPendingAnimeMetadata(tracker);

    tracker.handleMediaChange('/tmp/Little Witch Academia S02E06.mkv', 'Episode 6');
    await waitForPendingAnimeMetadata(tracker);

    const privateApi = tracker as unknown as {
      db: DatabaseSync;
    };
    const rows = privateApi.db
      .prepare(
        `
          SELECT
            v.source_path,
            v.anime_id,
            v.parsed_episode,
            a.canonical_title AS anime_title,
            a.anilist_id
          FROM imm_videos v
          LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
          WHERE v.source_path IN (?, ?)
          ORDER BY v.source_path
        `,
      )
      .all(
        '/tmp/Little Witch Academia S02E05.mkv',
        '/tmp/Little Witch Academia S02E06.mkv',
      ) as Array<{
      source_path: string | null;
      anime_id: number | null;
      parsed_episode: number | null;
      anime_title: string | null;
      anilist_id: number | null;
    }>;

    assert.equal(rows.length, 2);
    assert.ok(rows[0]?.anime_id);
    assert.equal(rows[0]?.anime_id, rows[1]?.anime_id);
    assert.deepEqual(
      rows.map((row) => ({
        sourcePath: row.source_path,
        parsedEpisode: row.parsed_episode,
        animeTitle: row.anime_title,
        anilistId: row.anilist_id,
      })),
      [
        {
          sourcePath: '/tmp/Little Witch Academia S02E05.mkv',
          parsedEpisode: 5,
          animeTitle: 'Little Witch Academia',
          anilistId: null,
        },
        {
          sourcePath: '/tmp/Little Witch Academia S02E06.mkv',
          parsedEpisode: 6,
          animeTitle: 'Little Witch Academia',
          anilistId: null,
        },
      ],
    );
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('applies configurable queue, flush, and retention policy', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({
      dbPath,
      policy: {
        batchSize: 10,
        flushIntervalMs: 250,
        queueCap: 1500,
        payloadCapBytes: 512,
        maintenanceIntervalMs: 2 * 60 * 60 * 1000,
        retention: {
          eventsDays: 14,
          telemetryDays: 45,
          sessionsDays: 60,
          dailyRollupsDays: 730,
          monthlyRollupsDays: 3650,
          vacuumIntervalDays: 14,
        },
      },
    });

    const privateApi = tracker as unknown as {
      batchSize: number;
      flushIntervalMs: number;
      queueCap: number;
      maxPayloadBytes: number;
      maintenanceIntervalMs: number;
      eventsRetentionMs: number;
      telemetryRetentionMs: number;
      sessionsRetentionMs: number;
      dailyRollupRetentionMs: number;
      monthlyRollupRetentionMs: number;
      vacuumIntervalMs: number;
    };

    assert.equal(privateApi.batchSize, 10);
    assert.equal(privateApi.flushIntervalMs, 250);
    assert.equal(privateApi.queueCap, 1500);
    assert.equal(privateApi.maxPayloadBytes, 512);
    assert.equal(privateApi.maintenanceIntervalMs, 7_200_000);
    assert.equal(privateApi.eventsRetentionMs, 14 * 86_400_000);
    assert.equal(privateApi.telemetryRetentionMs, 45 * 86_400_000);
    assert.equal(privateApi.sessionsRetentionMs, 60 * 86_400_000);
    assert.equal(privateApi.dailyRollupRetentionMs, 730 * 86_400_000);
    assert.equal(privateApi.monthlyRollupRetentionMs, 3650 * 86_400_000);
    assert.equal(privateApi.vacuumIntervalMs, 14 * 86_400_000);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('zero retention days disables prune checks while preserving rollups', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({
      dbPath,
      policy: {
        retention: {
          eventsDays: 0,
          telemetryDays: 0,
          sessionsDays: 0,
          dailyRollupsDays: 0,
          monthlyRollupsDays: 0,
          vacuumIntervalDays: 0,
        },
      },
    });

    const privateApi = tracker as unknown as {
      runMaintenance: () => void;
      db: DatabaseSync;
      eventsRetentionMs: number;
      telemetryRetentionMs: number;
      sessionsRetentionMs: number;
      dailyRollupRetentionMs: number;
      monthlyRollupRetentionMs: number;
      vacuumIntervalMs: number;
      lastVacuumMs: number;
    };

    assert.equal(privateApi.eventsRetentionMs, Number.POSITIVE_INFINITY);
    assert.equal(privateApi.telemetryRetentionMs, Number.POSITIVE_INFINITY);
    assert.equal(privateApi.sessionsRetentionMs, Number.POSITIVE_INFINITY);
    assert.equal(privateApi.dailyRollupRetentionMs, Number.POSITIVE_INFINITY);
    assert.equal(privateApi.monthlyRollupRetentionMs, Number.POSITIVE_INFINITY);
    assert.equal(privateApi.vacuumIntervalMs, Number.POSITIVE_INFINITY);
    assert.equal(privateApi.lastVacuumMs, 0);

    const nowMs = trackerNowMs();
    const oldMs = nowMs - 40 * 86_400_000;
    const olderMs = nowMs - 70 * 86_400_000;
    const insertedDailyRollupKeys = [1_000_001, 1_000_002];
    const insertedMonthlyRollupKeys = [202212, 202301];

    privateApi.db.exec(`
      INSERT INTO imm_videos (
        video_id,
        video_key,
        canonical_title,
        source_type,
        duration_ms,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        1,
        'local:/tmp/video.mkv',
        'Episode',
        1,
        0,
        ${olderMs},
        ${olderMs}
      )
    `);
    privateApi.db.exec(`
      INSERT INTO imm_sessions (
        session_id,
        session_uuid,
        video_id,
        started_at_ms,
        ended_at_ms,
        status,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
      (1, 'session-1', 1, ${olderMs}, ${olderMs + 1_000}, 2, ${olderMs}, ${olderMs}),
      (2, 'session-2', 1, ${oldMs}, ${oldMs + 1_000}, 2, ${oldMs}, ${oldMs})
    `);
    privateApi.db.exec(`
      INSERT INTO imm_session_events (
        session_id,
        ts_ms,
        event_type,
        segment_start_ms,
        segment_end_ms,
        created_date,
        last_update_date
      ) VALUES
      (1, ${olderMs}, 1, 0, 1, ${olderMs}, ${olderMs}),
      (2, ${oldMs}, 1, 2, 3, ${oldMs}, ${oldMs})
    `);
    privateApi.db.exec(`
      INSERT INTO imm_session_telemetry (
        session_id,
        sample_ms,
        total_watched_ms,
        active_watched_ms,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
      (1, ${olderMs}, 1000, 1000, ${olderMs}, ${olderMs}),
      (2, ${oldMs}, 2000, 1500, ${oldMs}, ${oldMs})
    `);
    privateApi.db.exec(`
      INSERT INTO imm_daily_rollups (
        rollup_day,
        video_id,
        total_sessions,
        total_active_min,
        total_lines_seen,
        total_tokens_seen,
        total_cards
      ) VALUES
      (${insertedDailyRollupKeys[0]}, 1, 1, 1, 1, 1, 1),
      (${insertedDailyRollupKeys[1]}, 1, 1, 1, 1, 1, 1)
    `);
    privateApi.db.exec(`
      INSERT INTO imm_monthly_rollups (
        rollup_month,
        video_id,
        total_sessions,
        total_active_min,
        total_lines_seen,
        total_tokens_seen,
        total_cards,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
      (${insertedMonthlyRollupKeys[0]}, 1, 1, 1, 1, 1, 1, ${olderMs}, ${olderMs}),
      (${insertedMonthlyRollupKeys[1]}, 1, 1, 1, 1, 1, 1, ${oldMs}, ${oldMs})
    `);

    privateApi.runMaintenance();

    const rawEvents = privateApi.db
      .prepare('SELECT COUNT(*) as total FROM imm_session_events WHERE session_id IN (1,2)')
      .get() as { total: number };
    const rawTelemetry = privateApi.db
      .prepare('SELECT COUNT(*) as total FROM imm_session_telemetry WHERE session_id IN (1,2)')
      .get() as { total: number };
    const endedSessions = privateApi.db
      .prepare('SELECT COUNT(*) as total FROM imm_sessions WHERE session_id IN (1,2)')
      .get() as { total: number };
    const dailyRollups = privateApi.db
      .prepare(
        'SELECT COUNT(*) as total FROM imm_daily_rollups WHERE video_id = 1 AND rollup_day IN (?, ?)',
      )
      .get(insertedDailyRollupKeys[0], insertedDailyRollupKeys[1]) as { total: number };
    const monthlyRollups = privateApi.db
      .prepare(
        'SELECT COUNT(*) as total FROM imm_monthly_rollups WHERE video_id = 1 AND rollup_month IN (?, ?)',
      )
      .get(insertedMonthlyRollupKeys[0], insertedMonthlyRollupKeys[1]) as { total: number };

    assert.equal(rawEvents.total, 2);
    assert.equal(rawTelemetry.total, 2);
    assert.equal(endedSessions.total, 2);
    assert.equal(dailyRollups.total, 2);
    assert.equal(monthlyRollups.total, 2);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('monthly rollups are grouped by calendar month', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      runRollupMaintenance: () => void;
    };

    const januaryStartedAtMs = 1_768_478_400_000;
    const februaryStartedAtMs = 1_771_156_800_000;

    privateApi.db.exec(`
      INSERT INTO imm_videos (
        video_id,
        video_key,
        canonical_title,
        source_type,
        duration_ms,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        1,
        'local:/tmp/video.mkv',
        'Episode',
        1,
        0,
        ${januaryStartedAtMs},
        ${januaryStartedAtMs}
      )
    `);

    privateApi.db.exec(`
      INSERT INTO imm_sessions (
        session_id,
        session_uuid,
        video_id,
        started_at_ms,
        status,
        CREATED_DATE,
        LAST_UPDATE_DATE,
        ended_at_ms
      ) VALUES (
        1,
        '11111111-1111-1111-1111-111111111111',
        1,
        ${januaryStartedAtMs},
        2,
        ${januaryStartedAtMs},
        ${januaryStartedAtMs},
        ${januaryStartedAtMs + 5000}
      )
    `);
    privateApi.db.exec(`
      INSERT INTO imm_session_telemetry (
        session_id,
        sample_ms,
        total_watched_ms,
        active_watched_ms,
        lines_seen,
        tokens_seen,
        cards_mined,
        lookup_count,
        lookup_hits,
        pause_count,
        pause_ms,
        seek_forward_count,
        seek_backward_count,
        media_buffer_events
      ) VALUES (
        1,
        ${januaryStartedAtMs + 1000},
        5000,
        5000,
        1,
        2,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      )
    `);

    privateApi.db.exec(`
      INSERT INTO imm_sessions (
        session_id,
        session_uuid,
        video_id,
        started_at_ms,
        status,
        CREATED_DATE,
        LAST_UPDATE_DATE,
        ended_at_ms
      ) VALUES (
        2,
        '22222222-2222-2222-2222-222222222222',
        1,
        ${februaryStartedAtMs},
        2,
        ${februaryStartedAtMs},
        ${februaryStartedAtMs},
        ${februaryStartedAtMs + 5000}
      )
    `);
    privateApi.db.exec(`
      INSERT INTO imm_session_telemetry (
        session_id,
        sample_ms,
        total_watched_ms,
        active_watched_ms,
        lines_seen,
        tokens_seen,
        cards_mined,
        lookup_count,
        lookup_hits,
        pause_count,
        pause_ms,
        seek_forward_count,
        seek_backward_count,
        media_buffer_events
      ) VALUES (
        2,
        ${februaryStartedAtMs + 1000},
        4000,
        4000,
        2,
        3,
        1,
        1,
        1,
        0,
        0,
        0,
        0,
        0
      )
    `);

    privateApi.db.exec(`
      INSERT INTO imm_monthly_rollups (
        rollup_month,
        video_id,
        total_sessions,
        total_active_min,
        total_lines_seen,
        total_tokens_seen,
        total_cards,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
      (202602, 1, 1, 1, 1, 1, 1, ${februaryStartedAtMs}, ${februaryStartedAtMs}),
      (202601, 1, 1, 1, 1, 1, 1, ${januaryStartedAtMs}, ${januaryStartedAtMs})
    `);

    const rows = await tracker.getMonthlyRollups(10);
    const videoRows = rows.filter((row) => row.videoId === 1);

    assert.equal(videoRows.length, 2);
    assert.equal(videoRows[0]!.rollupDayOrMonth, 202602);
    assert.equal(videoRows[1]!.rollupDayOrMonth, 202601);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('flushSingle reuses cached prepared statements', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;
  let originalPrepare: DatabaseSync['prepare'] | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      flushSingle: (write: {
        kind: 'telemetry' | 'event';
        sessionId: number;
        sampleMs: number;
        eventType?: number;
        lineIndex?: number | null;
        segmentStartMs?: number | null;
        segmentEndMs?: number | null;
        tokensDelta?: number;
        cardsDelta?: number;
        payloadJson?: string | null;
        totalWatchedMs?: number;
        activeWatchedMs?: number;
        linesSeen?: number;
        tokensSeen?: number;
        cardsMined?: number;
        lookupCount?: number;
        lookupHits?: number;
        yomitanLookupCount?: number;
        pauseCount?: number;
        pauseMs?: number;
        seekForwardCount?: number;
        seekBackwardCount?: number;
        mediaBufferEvents?: number;
      }) => void;
    };

    originalPrepare = privateApi.db.prepare;
    let prepareCalls = 0;
    privateApi.db.prepare = (...args: Parameters<DatabaseSync['prepare']>) => {
      prepareCalls += 1;
      return originalPrepare!.apply(privateApi.db, args);
    };
    const preparedRestore = originalPrepare;

    privateApi.db.exec(`
      INSERT INTO imm_videos (
        video_id,
        video_key,
        canonical_title,
        source_type,
        duration_ms,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        1,
        'local:/tmp/prepared.mkv',
        'Prepared',
        1,
        0,
        1000,
        1000
      )
    `);

    privateApi.db.exec(`
      INSERT INTO imm_sessions (
        session_id,
        session_uuid,
        video_id,
        started_at_ms,
        status,
        CREATED_DATE,
        LAST_UPDATE_DATE,
        ended_at_ms
      ) VALUES (
        1,
        '33333333-3333-3333-3333-333333333333',
        1,
        1000,
        2,
        1000,
        1000,
        2000
      )
    `);

    privateApi.flushSingle({
      kind: 'telemetry',
      sessionId: 1,
      sampleMs: 1500,
      totalWatchedMs: 1000,
      activeWatchedMs: 1000,
      linesSeen: 1,
      tokensSeen: 2,
      cardsMined: 0,
      lookupCount: 0,
      lookupHits: 0,
      yomitanLookupCount: 0,
      pauseCount: 0,
      pauseMs: 0,
      seekForwardCount: 0,
      seekBackwardCount: 0,
      mediaBufferEvents: 0,
    });

    privateApi.flushSingle({
      kind: 'event',
      sessionId: 1,
      sampleMs: 1600,
      eventType: 1,
      lineIndex: 1,
      segmentStartMs: 0,
      segmentEndMs: 1000,
      tokensDelta: 2,
      cardsDelta: 0,
      payloadJson: '{"event":"subtitle-line"}',
    });

    privateApi.db.prepare = preparedRestore;

    assert.equal(prepareCalls, 0);
  } finally {
    if (tracker && originalPrepare) {
      const privateApi = tracker as unknown as { db: DatabaseSync };
      privateApi.db.prepare = originalPrepare;
    }
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('reassignAnimeAnilist deduplicates cover blobs and getCoverArt remains compatible', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;
  const originalFetch = globalThis.fetch;
  const sharedCoverBlob = Buffer.from([1, 2, 3, 4, 5, 6, 7, 8]);

  try {
    globalThis.fetch = async () =>
      new Response(new Uint8Array(sharedCoverBlob), {
        status: 200,
        headers: { 'Content-Type': 'image/jpeg' },
      });
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as { db: DatabaseSync };

    privateApi.db.exec(`
      INSERT INTO imm_anime (
        anime_id,
        normalized_title_key,
        canonical_title,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        1,
        'little witch academia',
        'Little Witch Academia',
        1000,
        1000
      );
      INSERT INTO imm_videos (
        video_id,
        video_key,
        canonical_title,
        source_type,
        duration_ms,
        anime_id,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
      (
        1,
        'local:/tmp/lwa-1.mkv',
        'Little Witch Academia S01E01',
        1,
        0,
        1,
        1000,
        1000
      ),
      (
        2,
        'local:/tmp/lwa-2.mkv',
        'Little Witch Academia S01E02',
        1,
        0,
        1,
        1000,
        1000
      );
    `);

    await tracker.reassignAnimeAnilist(1, {
      anilistId: 33489,
      titleRomaji: 'Little Witch Academia',
      titleEnglish: 'Little Witch Academia',
      episodesTotal: 25,
      coverUrl: 'https://example.com/lwa.jpg',
    });

    const blobRows = privateApi.db
      .prepare('SELECT blob_hash AS blobHash, cover_blob AS coverBlob FROM imm_cover_art_blobs')
      .all() as Array<{ blobHash: string; coverBlob: Buffer }>;
    const mediaRows = privateApi.db
      .prepare(
        `
          SELECT
            video_id AS videoId,
            cover_blob AS coverBlob,
            cover_blob_hash AS coverBlobHash
          FROM imm_media_art
          ORDER BY video_id ASC
        `,
      )
      .all() as Array<{
      videoId: number;
      coverBlob: Buffer | null;
      coverBlobHash: string | null;
    }>;

    assert.equal(blobRows.length, 1);
    assert.deepEqual(new Uint8Array(blobRows[0]!.coverBlob), new Uint8Array(sharedCoverBlob));
    assert.equal(mediaRows.length, 2);
    assert.equal(typeof mediaRows[0]?.coverBlobHash, 'string');
    assert.equal(mediaRows[0]?.coverBlobHash, mediaRows[1]?.coverBlobHash);

    const resolvedCover = await tracker.getCoverArt(2);
    assert.ok(resolvedCover?.coverBlob);
    assert.deepEqual(
      new Uint8Array(resolvedCover?.coverBlob ?? Buffer.alloc(0)),
      new Uint8Array(sharedCoverBlob),
    );
  } finally {
    globalThis.fetch = originalFetch;
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('reassignAnimeAnilist replaces stale cover blobs when the AniList cover changes', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;
  const originalFetch = globalThis.fetch;
  const initialCoverBlob = Buffer.from([1, 2, 3, 4]);
  const replacementCoverBlob = Buffer.from([9, 8, 7, 6]);
  let fetchCallCount = 0;

  try {
    globalThis.fetch = async () => {
      fetchCallCount += 1;
      const blob = fetchCallCount === 1 ? initialCoverBlob : replacementCoverBlob;
      return new Response(new Uint8Array(blob), {
        status: 200,
        headers: { 'Content-Type': 'image/jpeg' },
      });
    };
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as { db: DatabaseSync };

    privateApi.db.exec(`
      INSERT INTO imm_anime (
        anime_id,
        normalized_title_key,
        canonical_title,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        1,
        'little witch academia',
        'Little Witch Academia',
        1000,
        1000
      );
      INSERT INTO imm_videos (
        video_id,
        video_key,
        canonical_title,
        source_type,
        duration_ms,
        anime_id,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
      (
        1,
        'local:/tmp/lwa-1.mkv',
        'Little Witch Academia S01E01',
        1,
        0,
        1,
        1000,
        1000
      ),
      (
        2,
        'local:/tmp/lwa-2.mkv',
        'Little Witch Academia S01E02',
        1,
        0,
        1,
        1000,
        1000
      );
    `);

    await tracker.reassignAnimeAnilist(1, {
      anilistId: 33489,
      titleRomaji: 'Little Witch Academia',
      coverUrl: 'https://example.com/lwa-old.jpg',
    });

    await tracker.reassignAnimeAnilist(1, {
      anilistId: 100526,
      titleRomaji: 'Otome Game Sekai wa Mob ni Kibishii Sekai desu',
      coverUrl: 'https://example.com/mobseka-new.jpg',
    });

    const mediaRows = privateApi.db
      .prepare(
        `
          SELECT
            video_id AS videoId,
            anilist_id AS anilistId,
            cover_url AS coverUrl,
            cover_blob_hash AS coverBlobHash
          FROM imm_media_art
          ORDER BY video_id ASC
        `,
      )
      .all() as Array<{
      videoId: number;
      anilistId: number | null;
      coverUrl: string | null;
      coverBlobHash: string | null;
    }>;
    const blobRows = privateApi.db
      .prepare('SELECT blob_hash AS blobHash, cover_blob AS coverBlob FROM imm_cover_art_blobs')
      .all() as Array<{ blobHash: string; coverBlob: Buffer }>;
    const resolvedCover = await tracker.getAnimeCoverArt(1);

    assert.equal(fetchCallCount, 2);
    assert.equal(mediaRows.length, 2);
    assert.equal(mediaRows[0]?.anilistId, 100526);
    assert.equal(mediaRows[0]?.coverUrl, 'https://example.com/mobseka-new.jpg');
    assert.equal(mediaRows[0]?.coverBlobHash, mediaRows[1]?.coverBlobHash);
    assert.equal(blobRows.length, 1);
    assert.deepEqual(
      new Uint8Array(blobRows[0]?.coverBlob ?? Buffer.alloc(0)),
      new Uint8Array(replacementCoverBlob),
    );
    assert.deepEqual(
      new Uint8Array(resolvedCover?.coverBlob ?? Buffer.alloc(0)),
      new Uint8Array(replacementCoverBlob),
    );
  } finally {
    globalThis.fetch = originalFetch;
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('reassignAnimeAnilist preserves existing description when description is omitted', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as { db: DatabaseSync };

    privateApi.db.exec(`
      INSERT INTO imm_anime (
        anime_id,
        normalized_title_key,
        canonical_title,
        description,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        1,
        'little witch academia',
        'Little Witch Academia',
        'Original description',
        1000,
        1000
      );
    `);

    await tracker.reassignAnimeAnilist(1, {
      anilistId: 33489,
      titleRomaji: 'Little Witch Academia',
    });

    const row = privateApi.db
      .prepare('SELECT anilist_id AS anilistId, description FROM imm_anime WHERE anime_id = ?')
      .get(1) as { anilistId: number | null; description: string | null } | null;

    assert.equal(row?.anilistId, 33489);
    assert.equal(row?.description, 'Original description');
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('handleMediaChange stores youtube metadata for new youtube sessions', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;
  const originalFetch = globalThis.fetch;
  const originalPath = process.env.PATH;
  let fakeBinDir: string | null = null;

  try {
    fakeBinDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-yt-dlp-bin-'));
    const ytDlpOutput =
      '{"id":"abc123","title":"Video Name","webpage_url":"https://www.youtube.com/watch?v=abc123","thumbnail":"https://i.ytimg.com/vi/abc123/hqdefault.jpg","channel_id":"UCcreator123","channel":"Creator Name","channel_url":"https://www.youtube.com/channel/UCcreator123","uploader_id":"@creator","uploader_url":"https://www.youtube.com/@creator","description":"Video description","channel_follower_count":12345,"thumbnails":[{"url":"https://i.ytimg.com/vi/abc123/hqdefault.jpg"},{"url":"https://yt3.googleusercontent.com/channel-avatar=s88"}]}';
    if (process.platform === 'win32') {
      const outputPath = path.join(fakeBinDir, 'output.json');
      fs.writeFileSync(outputPath, ytDlpOutput, 'utf8');
      fs.writeFileSync(
        path.join(fakeBinDir, 'yt-dlp.cmd'),
        '@echo off\r\ntype "%~dp0output.json"\r\n',
        'utf8',
      );
    } else {
      const scriptPath = path.join(fakeBinDir, 'yt-dlp');
      fs.writeFileSync(
        scriptPath,
        `#!/bin/sh
printf '%s\n' '${ytDlpOutput}'
`,
        { mode: 0o755 },
      );
    }
    process.env.PATH = `${fakeBinDir}${path.delimiter}${originalPath ?? ''}`;

    globalThis.fetch = async (input) => {
      const url = String(input);
      if (url.includes('/oembed')) {
        return new Response(
          JSON.stringify({
            thumbnail_url: 'https://i.ytimg.com/vi/abc123/hqdefault.jpg',
          }),
          { status: 200, headers: { 'Content-Type': 'application/json' } },
        );
      }
      return new Response(new Uint8Array([1, 2, 3]), {
        status: 200,
        headers: { 'Content-Type': 'image/jpeg' },
      });
    };

    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('https://www.youtube.com/watch?v=abc123', 'Player Title');
    const privateApi = tracker as unknown as { db: DatabaseSync };
    await waitForCondition(() => {
      const stored = privateApi.db
        .prepare("SELECT 1 AS ready FROM imm_youtube_videos WHERE youtube_video_id = 'abc123'")
        .get() as { ready: number } | null;
      return stored?.ready === 1;
    }, 5_000);
    const row = privateApi.db
      .prepare(
        `
          SELECT
            youtube_video_id AS youtubeVideoId,
            video_url AS videoUrl,
            video_title AS videoTitle,
            video_thumbnail_url AS videoThumbnailUrl,
            channel_id AS channelId,
            channel_name AS channelName,
            channel_url AS channelUrl,
            channel_thumbnail_url AS channelThumbnailUrl,
            uploader_id AS uploaderId,
            uploader_url AS uploaderUrl,
            description AS description
          FROM imm_youtube_videos
        `,
      )
      .get() as {
      youtubeVideoId: string;
      videoUrl: string;
      videoTitle: string;
      videoThumbnailUrl: string;
      channelId: string;
      channelName: string;
      channelUrl: string;
      channelThumbnailUrl: string;
      uploaderId: string;
      uploaderUrl: string;
      description: string;
    } | null;
    const videoRow = privateApi.db
      .prepare(
        `
          SELECT canonical_title AS canonicalTitle
          FROM imm_videos
          WHERE video_id = 1
        `,
      )
      .get() as { canonicalTitle: string } | null;
    const animeRow = privateApi.db
      .prepare(
        `
          SELECT
            a.canonical_title AS canonicalTitle,
            v.parsed_title AS parsedTitle,
            v.parser_source AS parserSource
          FROM imm_videos v
          JOIN imm_anime a ON a.anime_id = v.anime_id
          WHERE v.video_id = 1
        `,
      )
      .get() as {
      canonicalTitle: string;
      parsedTitle: string | null;
      parserSource: string | null;
    } | null;

    assert.ok(row);
    assert.ok(videoRow);
    assert.equal(row.youtubeVideoId, 'abc123');
    assert.equal(row.videoUrl, 'https://www.youtube.com/watch?v=abc123');
    assert.equal(row.videoTitle, 'Video Name');
    assert.equal(row.videoThumbnailUrl, 'https://i.ytimg.com/vi/abc123/hqdefault.jpg');
    assert.equal(row.channelId, 'UCcreator123');
    assert.equal(row.channelName, 'Creator Name');
    assert.equal(row.channelUrl, 'https://www.youtube.com/channel/UCcreator123');
    assert.equal(row.channelThumbnailUrl, 'https://yt3.googleusercontent.com/channel-avatar=s88');
    assert.equal(row.uploaderId, '@creator');
    assert.equal(row.uploaderUrl, 'https://www.youtube.com/@creator');
    assert.equal(row.description, 'Video description');
    assert.equal(videoRow.canonicalTitle, 'Video Name');
    assert.equal(animeRow?.canonicalTitle, 'Creator Name');
    assert.equal(animeRow?.parsedTitle, 'Creator Name');
    assert.equal(animeRow?.parserSource, 'youtube');
  } finally {
    process.env.PATH = originalPath;
    globalThis.fetch = originalFetch;
    tracker?.destroy();
    cleanupDbPath(dbPath);
    if (fakeBinDir) {
      fs.rmSync(fakeBinDir, { recursive: true, force: true });
    }
  }
});

test('getMediaLibrary lazily backfills missing youtube metadata for existing rows', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;
  const originalPath = process.env.PATH;
  let fakeBinDir: string | null = null;

  try {
    fakeBinDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-yt-dlp-bin-'));
    const ytDlpOutput =
      '{"id":"backfill123","title":"Backfilled Video Title","webpage_url":"https://www.youtube.com/watch?v=backfill123","thumbnail":"https://i.ytimg.com/vi/backfill123/hqdefault.jpg","channel_id":"UCbackfill123","channel":"Backfill Creator","channel_url":"https://www.youtube.com/channel/UCbackfill123","uploader_id":"@backfill","uploader_url":"https://www.youtube.com/@backfill","description":"Backfilled description","thumbnails":[{"url":"https://i.ytimg.com/vi/backfill123/hqdefault.jpg"},{"url":"https://yt3.googleusercontent.com/backfill-avatar=s88"}]}';
    if (process.platform === 'win32') {
      const outputPath = path.join(fakeBinDir, 'output.json');
      fs.writeFileSync(outputPath, ytDlpOutput, 'utf8');
      fs.writeFileSync(
        path.join(fakeBinDir, 'yt-dlp.cmd'),
        '@echo off\r\ntype "%~dp0output.json"\r\n',
        'utf8',
      );
    } else {
      const scriptPath = path.join(fakeBinDir, 'yt-dlp');
      fs.writeFileSync(
        scriptPath,
        `#!/bin/sh
printf '%s\n' '${ytDlpOutput}'
`,
        { mode: 0o755 },
      );
    }
    process.env.PATH = `${fakeBinDir}${path.delimiter}${originalPath ?? ''}`;

    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as { db: DatabaseSync };
    const nowMs = trackerNowMs();

    privateApi.db
      .prepare(
        `
          INSERT INTO imm_videos (
            video_key,
            canonical_title,
            source_type,
            source_path,
            source_url,
            duration_ms,
            file_size_bytes,
            codec_id,
            container_id,
            width_px,
            height_px,
            fps_x100,
            bitrate_kbps,
            audio_codec_id,
            hash_sha256,
            screenshot_path,
            metadata_json,
            CREATED_DATE,
            LAST_UPDATE_DATE
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `,
      )
      .run(
        'remote:https://www.youtube.com/watch?v=backfill123',
        'watch?v=backfill123',
        2,
        null,
        'https://www.youtube.com/watch?v=backfill123',
        0,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        nowMs,
        nowMs,
      );
    privateApi.db
      .prepare(
        `
          INSERT INTO imm_lifetime_media (
            video_id,
            total_sessions,
            total_active_ms,
            total_cards,
            total_lines_seen,
            total_tokens_seen,
            completed,
            first_watched_ms,
            last_watched_ms,
            CREATED_DATE,
            LAST_UPDATE_DATE
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `,
      )
      .run(1, 1, 5_000, 0, 0, 50, 0, nowMs, nowMs, nowMs, nowMs);

    const before = await tracker.getMediaLibrary();
    assert.equal(before[0]?.channelName ?? null, null);

    await waitForCondition(() => {
      const row = privateApi.db
        .prepare(
          `
            SELECT
              video_title AS videoTitle,
              channel_name AS channelName,
              channel_thumbnail_url AS channelThumbnailUrl
            FROM imm_youtube_videos
            WHERE video_id = 1
          `,
        )
        .get() as {
        videoTitle: string | null;
        channelName: string | null;
        channelThumbnailUrl: string | null;
      } | null;
      return (
        row?.videoTitle === 'Backfilled Video Title' &&
        row.channelName === 'Backfill Creator' &&
        row.channelThumbnailUrl === 'https://yt3.googleusercontent.com/backfill-avatar=s88'
      );
    }, 5_000);

    const after = await tracker.getMediaLibrary();
    assert.equal(after[0]?.videoTitle, 'Backfilled Video Title');
    assert.equal(after[0]?.channelName, 'Backfill Creator');
    assert.equal(
      after[0]?.channelThumbnailUrl,
      'https://yt3.googleusercontent.com/backfill-avatar=s88',
    );
  } finally {
    process.env.PATH = originalPath;
    tracker?.destroy();
    cleanupDbPath(dbPath);
    if (fakeBinDir) {
      fs.rmSync(fakeBinDir, { recursive: true, force: true });
    }
  }
});

test('getAnimeLibrary lazily relinks youtube rows to channel groupings', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as { db: DatabaseSync };
    const nowMs = trackerNowMs();

    privateApi.db.exec(`
      INSERT INTO imm_anime (
        anime_id,
        normalized_title_key,
        canonical_title,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
        (1, 'watch v first', 'watch?v first', ${nowMs}, ${nowMs}),
        (2, 'watch v second', 'watch?v second', ${nowMs}, ${nowMs});

      INSERT INTO imm_videos (
        video_id,
        anime_id,
        video_key,
        canonical_title,
        parsed_title,
        parser_source,
        source_type,
        source_path,
        source_url,
        duration_ms,
        file_size_bytes,
        codec_id,
        container_id,
        width_px,
        height_px,
        fps_x100,
        bitrate_kbps,
        audio_codec_id,
        hash_sha256,
        screenshot_path,
        metadata_json,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
        (
          1,
          1,
          'remote:https://www.youtube.com/watch?v=first',
          'watch?v first',
          'watch?v first',
          'fallback',
          2,
          NULL,
          'https://www.youtube.com/watch?v=first',
          0,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          ${nowMs},
          ${nowMs}
        ),
        (
          2,
          2,
          'remote:https://www.youtube.com/watch?v=second',
          'watch?v second',
          'watch?v second',
          'fallback',
          2,
          NULL,
          'https://www.youtube.com/watch?v=second',
          0,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          NULL,
          ${nowMs},
          ${nowMs}
        );

      INSERT INTO imm_youtube_videos (
        video_id,
        youtube_video_id,
        video_url,
        video_title,
        video_thumbnail_url,
        channel_id,
        channel_name,
        channel_url,
        channel_thumbnail_url,
        uploader_id,
        uploader_url,
        description,
        metadata_json,
        fetched_at_ms,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
        (
          1,
          'first',
          'https://www.youtube.com/watch?v=first',
          'First Video',
          'https://i.ytimg.com/vi/first/hqdefault.jpg',
          'UCchannel1',
          'Shared Channel',
          'https://www.youtube.com/channel/UCchannel1',
          'https://yt3.googleusercontent.com/shared=s88',
          '@shared',
          'https://www.youtube.com/@shared',
          NULL,
          '{}',
          ${nowMs},
          ${nowMs},
          ${nowMs}
        ),
        (
          2,
          'second',
          'https://www.youtube.com/watch?v=second',
          'Second Video',
          'https://i.ytimg.com/vi/second/hqdefault.jpg',
          'UCchannel1',
          'Shared Channel',
          'https://www.youtube.com/channel/UCchannel1',
          'https://yt3.googleusercontent.com/shared=s88',
          '@shared',
          'https://www.youtube.com/@shared',
          NULL,
          '{}',
          ${nowMs},
          ${nowMs},
          ${nowMs}
        );

      INSERT INTO imm_sessions (
        session_id,
        session_uuid,
        video_id,
        started_at_ms,
        ended_at_ms,
        status,
        total_watched_ms,
        active_watched_ms,
        lines_seen,
        tokens_seen,
        cards_mined,
        lookup_count,
        lookup_hits,
        yomitan_lookup_count,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
        (
          1,
          'session-youtube-1',
          1,
          ${nowMs - 70000},
          ${nowMs - 10000},
          2,
          65000,
          60000,
          0,
          100,
          0,
          0,
          0,
          0,
          ${nowMs},
          ${nowMs}
        ),
        (
          2,
          'session-youtube-2',
          2,
          ${nowMs - 50000},
          ${nowMs - 5000},
          2,
          35000,
          30000,
          0,
          50,
          0,
          0,
          0,
          0,
          ${nowMs},
          ${nowMs}
        );

      INSERT INTO imm_lifetime_anime (
        anime_id,
        total_sessions,
        total_active_ms,
        total_cards,
        total_lines_seen,
        total_tokens_seen,
        episodes_started,
        episodes_completed,
        first_watched_ms,
        last_watched_ms,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
        (1, 1, 60000, 0, 0, 100, 1, 0, ${nowMs}, ${nowMs}, ${nowMs}, ${nowMs}),
        (2, 1, 30000, 0, 0, 50, 1, 0, ${nowMs}, ${nowMs}, ${nowMs}, ${nowMs});

      INSERT INTO imm_lifetime_media (
        video_id,
        total_sessions,
        total_active_ms,
        total_cards,
        total_lines_seen,
        total_tokens_seen,
        completed,
        first_watched_ms,
        last_watched_ms,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES
        (1, 1, 60000, 0, 0, 100, 0, ${nowMs}, ${nowMs}, ${nowMs}, ${nowMs}),
        (2, 1, 30000, 0, 0, 50, 0, ${nowMs}, ${nowMs}, ${nowMs}, ${nowMs});
    `);

    const rows = await tracker.getAnimeLibrary();
    const sharedRows = rows.filter((row) => row.canonicalTitle === 'Shared Channel');

    assert.equal(sharedRows.length, 1);
    assert.equal(sharedRows[0]?.episodeCount, 2);

    const relinked = privateApi.db
      .prepare(
        `
          SELECT a.canonical_title AS canonicalTitle, COUNT(*) AS total
          FROM imm_videos v
          JOIN imm_anime a ON a.anime_id = v.anime_id
          GROUP BY a.anime_id, a.canonical_title
          ORDER BY total DESC, a.anime_id ASC
        `,
      )
      .all() as Array<{ canonicalTitle: string; total: number }>;

    assert.equal(relinked[0]?.canonicalTitle, 'Shared Channel');
    assert.equal(relinked[0]?.total, 2);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('reassignAnimeAnilist clears description when description is explicitly null', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as { db: DatabaseSync };

    privateApi.db.exec(`
      INSERT INTO imm_anime (
        anime_id,
        normalized_title_key,
        canonical_title,
        description,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        1,
        'little witch academia',
        'Little Witch Academia',
        'Original description',
        1000,
        1000
      );
    `);

    await tracker.reassignAnimeAnilist(1, {
      anilistId: 33489,
      description: null,
    });

    const row = privateApi.db
      .prepare('SELECT description FROM imm_anime WHERE anime_id = ?')
      .get(1) as { description: string | null } | null;

    assert.equal(row?.description, null);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('ensureCoverArt returns false when fetcher reports success without storing art', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;
  let fetchCalls = 0;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as { db: DatabaseSync };

    privateApi.db.exec(`
      INSERT INTO imm_videos (
        video_id,
        video_key,
        canonical_title,
        source_type,
        duration_ms,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        1,
        'local:/tmp/lwa-1.mkv',
        'Little Witch Academia S01E01',
        1,
        0,
        1000,
        1000
      );
      INSERT INTO imm_lifetime_media (
        video_id,
        total_sessions,
        total_active_ms,
        total_cards,
        total_tokens_seen,
        total_lines_seen,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        1,
        0,
        0,
        0,
        0,
        0,
        1000,
        1000
      );
    `);

    tracker.setCoverArtFetcher({
      fetchIfMissing: async () => {
        fetchCalls += 1;
        return true;
      },
    });

    const storedBefore = await tracker.getCoverArt(1);
    assert.equal(storedBefore?.coverBlob ?? null, null);

    const result = await tracker.ensureCoverArt(1);

    assert.equal(fetchCalls, 1);
    assert.equal(result, false);
    assert.equal((await tracker.getCoverArt(1))?.coverBlob ?? null, null);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('markActiveVideoWatched marks current session video as watched', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/test-mark-active.mkv', 'Test Mark Active');
    await waitForPendingAnimeMetadata(tracker);

    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      sessionState: { videoId: number; markedWatched: boolean } | null;
    };
    const videoId = privateApi.sessionState?.videoId;
    assert.ok(videoId);

    const result = await tracker.markActiveVideoWatched();
    assert.equal(result, true);
    assert.equal(privateApi.sessionState?.markedWatched, true);

    const row = privateApi.db
      .prepare('SELECT watched FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { watched: number } | null;
    assert.equal(row?.watched, 1);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('markActiveVideoWatched returns false when no active session', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const result = await tracker.markActiveVideoWatched();
    assert.equal(result, false);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});
