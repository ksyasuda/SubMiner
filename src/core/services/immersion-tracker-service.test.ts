import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { toMonthKey } from './immersion-tracker/maintenance';
import { enqueueWrite } from './immersion-tracker/queue';
import { Database, type DatabaseSync } from './immersion-tracker/sqlite';
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
  assert.equal(toMonthKey(Date.UTC(2026, 0, 31, 23, 59, 59, 999)), 202601);
  assert.equal(toMonthKey(Date.UTC(2026, 1, 1, 0, 0, 0, 0)), 202602);
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
      ended_at_ms: number | null;
    } | null;
    const telemetryCountRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_session_telemetry')
      .get() as { total: number };
    db.close();

    assert.ok(sessionRow);
    assert.ok(Number(sessionRow?.ended_at_ms ?? 0) > 0);
    assert.ok(Number(telemetryCountRow.total) >= 2);
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
    tracker.recordSubtitleLine('alpha beta', 0, 1.2);
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
        `SELECT lines_seen, words_seen, tokens_seen, cards_mined
         FROM imm_session_telemetry
         ORDER BY sample_ms DESC, telemetry_id DESC
         LIMIT 1`,
      )
      .get() as {
      lines_seen: number;
      words_seen: number;
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
    assert.ok(Number(telemetryRow?.words_seen ?? 0) >= 2);
    assert.ok(Number(telemetryRow?.tokens_seen ?? 0) >= 2);
    assert.ok(Number(telemetryRow?.cards_mined ?? 0) >= 2);
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
      .all('/tmp/Little Witch Academia S02E05.mkv', '/tmp/Little Witch Academia S02E06.mkv') as
      Array<{
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
    assert.equal(privateApi.dailyRollupRetentionMs, 730 * 86_400_000);
    assert.equal(privateApi.monthlyRollupRetentionMs, 3650 * 86_400_000);
    assert.equal(privateApi.vacuumIntervalMs, 14 * 86_400_000);
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

    const januaryStartedAtMs = Date.UTC(2026, 0, 31, 23, 59, 59, 0);
    const februaryStartedAtMs = Date.UTC(2026, 1, 1, 0, 0, 1, 0);

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
        words_seen,
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
        words_seen,
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

    privateApi.runRollupMaintenance();

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
        wordsDelta?: number;
        cardsDelta?: number;
        payloadJson?: string | null;
        totalWatchedMs?: number;
        activeWatchedMs?: number;
        linesSeen?: number;
        wordsSeen?: number;
        tokensSeen?: number;
        cardsMined?: number;
        lookupCount?: number;
        lookupHits?: number;
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
      wordsSeen: 2,
      tokensSeen: 2,
      cardsMined: 0,
      lookupCount: 0,
      lookupHits: 0,
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
      wordsDelta: 2,
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
