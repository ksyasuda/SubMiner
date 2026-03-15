import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Database } from '../sqlite.js';
import {
  createTrackerPreparedStatements,
  ensureSchema,
  getOrCreateAnimeRecord,
  getOrCreateVideoRecord,
  linkVideoToAnimeRecord,
} from '../storage.js';
import { startSessionRecord } from '../session.js';
import {
  cleanupVocabularyStats,
  getAnimeDetail,
  getAnimeEpisodes,
  getAnimeLibrary,
  getKanjiOccurrences,
  getSessionSummaries,
  getVocabularyStats,
  getKanjiStats,
  getSessionEvents,
  getWordOccurrences,
} from '../query.js';
import { SOURCE_TYPE_LOCAL, EVENT_SUBTITLE_LINE } from '../types.js';

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-imm-query-test-'));
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
  let lastError: NodeJS.ErrnoException | null = null;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      fs.rmSync(dir, { recursive: true, force: true });
      return;
    } catch (error) {
      const err = error as NodeJS.ErrnoException;
      lastError = err;
      if (process.platform !== 'win32' || err.code !== 'EBUSY') {
        throw error;
      }
      bunRuntime.Bun?.gc?.(true);
      Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 25);
    }
  }
  if (lastError) {
    throw lastError;
  }
}

test('getSessionSummaries returns sessionId and canonicalTitle', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/query-test.mkv', {
      canonicalTitle: 'Query Test Episode',
      sourcePath: '/tmp/query-test.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const startedAtMs = 1_000_000;
    const { sessionId } = startSessionRecord(db, videoId, startedAtMs);

    stmts.telemetryInsertStmt.run(
      sessionId,
      startedAtMs + 1_000,
      3_000,
      2_500,
      5,
      10,
      10,
      1,
      2,
      1,
      0,
      0,
      0,
      0,
      0,
      startedAtMs + 1_000,
      startedAtMs + 1_000,
    );

    const rows = getSessionSummaries(db, 10);

    assert.ok(rows.length >= 1);
    const row = rows.find((r) => r.sessionId === sessionId);
    assert.ok(row, 'expected to find a row for the created session');
    assert.equal(typeof row.sessionId, 'number');
    assert.equal(row.sessionId, sessionId);
    assert.equal(row.canonicalTitle, 'Query Test Episode');
    assert.equal(row.videoId, videoId);
    assert.equal(row.linesSeen, 5);
    assert.equal(row.totalWatchedMs, 3_000);
    assert.equal(row.activeWatchedMs, 2_500);
    assert.equal(row.wordsSeen, 10);
    assert.equal(row.tokensSeen, 10);
    assert.equal(row.lookupCount, 2);
    assert.equal(row.lookupHits, 1);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getSessionSummaries with no telemetry returns zero aggregates', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/no-telemetry.mkv', {
      canonicalTitle: 'No Telemetry',
      sourcePath: '/tmp/no-telemetry.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const { sessionId } = startSessionRecord(db, videoId, 3_000_000);

    const rows = getSessionSummaries(db, 10);
    const row = rows.find((r) => r.sessionId === sessionId);
    assert.ok(row, 'expected to find the session with no telemetry');
    assert.equal(row.canonicalTitle, 'No Telemetry');
    assert.equal(row.totalWatchedMs, 0);
    assert.equal(row.activeWatchedMs, 0);
    assert.equal(row.linesSeen, 0);
    assert.equal(row.wordsSeen, 0);
    assert.equal(row.tokensSeen, 0);
    assert.equal(row.lookupCount, 0);
    assert.equal(row.lookupHits, 0);
    assert.equal(row.cardsMined, 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getVocabularyStats returns rows ordered by frequency descending', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    // Insert words with the highest-frequency entry inserted after another word
    stmts.wordUpsertStmt.run('犬', '犬', 'いぬ', 'noun', '名詞', '一般', '', 1_500, 1_500);
    stmts.wordUpsertStmt.run('猫', '猫', 'ねこ', 'noun', '名詞', '一般', '', 1_000, 2_000);
    stmts.wordUpsertStmt.run('猫', '猫', 'ねこ', 'noun', '名詞', '一般', '', 1_000, 3_000);

    const rows = getVocabularyStats(db, 10);

    assert.equal(rows.length, 2);
    assert.equal(rows[0]?.headword, '猫');
    assert.equal(rows[1]?.headword, '犬');
    assert.equal(rows[0]?.frequency, 2);
    assert.equal(rows[1]?.frequency, 1);

    assert.ok(rows.length >= 2);
    // First row should be 猫 (frequency 2)
    const nekRow = rows.find((r) => r.headword === '猫');
    const inuRow = rows.find((r) => r.headword === '犬');
    assert.ok(nekRow, 'expected 猫 row');
    assert.ok(inuRow, 'expected 犬 row');
    assert.equal(nekRow.headword, '猫');
    assert.equal(nekRow.word, '猫');
    assert.equal(nekRow.reading, 'ねこ');
    assert.equal(nekRow.frequency, 2);
    assert.equal(typeof nekRow.firstSeen, 'number');
    assert.equal(typeof nekRow.lastSeen, 'number');
    // Higher frequency should come first
    const nekIdx = rows.indexOf(nekRow);
    const inuIdx = rows.indexOf(inuRow);
    assert.ok(nekIdx < inuIdx, 'higher frequency word should appear first');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getVocabularyStats returns empty array when no words exist', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const rows = getVocabularyStats(db, 10);
    assert.deepEqual(rows, []);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('cleanupVocabularyStats repairs stored POS metadata and removes excluded imm_words rows', async () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    db.prepare(
      `INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    ).run('猫', '猫', 'ねこ', 'noun', '名詞', '一般', '', 1_000, 1_500, 3);
    db.prepare(
      `INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    ).run('知っている', '知っている', '', 'other', '動詞', '自立', '', 1_025, 1_525, 4);
    db.prepare(
      `INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    ).run('は', 'は', 'は', 'particle', '助詞', '係助詞', '', 1_100, 1_600, 9);
    db.prepare(
      `INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    ).run('旧', '旧', '', '', '', '', '', 900, 950, 1);
    db.prepare(
      `INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    ).run('未解決', '未解決', '', '', '', '', '', 901, 951, 1);

    const result = await cleanupVocabularyStats(db, {
      resolveLegacyPos: async (row) => {
        if (row.headword === '旧') {
          return {
            partOfSpeech: 'noun',
            headword: '旧',
            reading: 'きゅう',
            pos1: '名詞',
            pos2: '一般',
            pos3: '',
          };
        }
        if (row.headword === '知っている') {
          return {
            partOfSpeech: 'verb',
            headword: '知る',
            reading: 'しっている',
            pos1: '動詞',
            pos2: '自立',
            pos3: '',
          };
        }
        return null;
      },
    });
    const rows = getVocabularyStats(db, 10);
    const repairedRows = db
      .prepare(
        `SELECT headword, word, reading, part_of_speech, pos1, pos2
         FROM imm_words
         ORDER BY headword ASC, word ASC`,
      )
      .all() as Array<{
      headword: string;
      word: string;
      reading: string;
      part_of_speech: string;
      pos1: string;
      pos2: string;
    }>;

    assert.deepEqual(result, { scanned: 5, kept: 3, deleted: 2, repaired: 2 });
    assert.deepEqual(
      rows.map((row) => ({ headword: row.headword, frequency: row.frequency })),
      [
        { headword: '知る', frequency: 4 },
        { headword: '猫', frequency: 3 },
        { headword: '旧', frequency: 1 },
      ],
    );
    assert.deepEqual(
      repairedRows,
      [
        {
          headword: '旧',
          word: '旧',
          reading: 'きゅう',
          part_of_speech: 'noun',
          pos1: '名詞',
          pos2: '一般',
        },
        {
          headword: '猫',
          word: '猫',
          reading: 'ねこ',
          part_of_speech: 'noun',
          pos1: '名詞',
          pos2: '一般',
        },
        {
          headword: '知る',
          word: '知っている',
          reading: 'しっている',
          part_of_speech: 'verb',
          pos1: '動詞',
          pos2: '自立',
        },
      ],
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('cleanupVocabularyStats merges repaired duplicates instead of violating the imm_words unique key', async () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/cleanup-merge.mkv', {
      canonicalTitle: 'Cleanup Merge',
      sourcePath: '/tmp/cleanup-merge.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const { sessionId } = startSessionRecord(db, videoId, 2_000_000);
    const duplicateResult = db
      .prepare(
        `INSERT INTO imm_words (
          headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run('知る', '知っている', 'しっている', 'verb', '動詞', '自立', '', 2_000, 2_500, 3);
    const legacyResult = db
      .prepare(
        `INSERT INTO imm_words (
          headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run('知っている', '知っている', '', 'other', '動詞', '自立', '', 1_000, 3_000, 4);
    const lineResult = db
      .prepare(
        `INSERT INTO imm_subtitle_lines (
          session_id, event_id, video_id, anime_id, line_index, segment_start_ms, segment_end_ms, text, CREATED_DATE, LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run(sessionId, null, videoId, null, 1, 0, 1000, '知っている', 2_000, 2_000);
    const lineId = Number(lineResult.lastInsertRowid);
    const duplicateId = Number(duplicateResult.lastInsertRowid);
    const legacyId = Number(legacyResult.lastInsertRowid);
    db.prepare(
      `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(lineId, duplicateId, 2);
    db.prepare(
      `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(lineId, legacyId, 1);

    const result = await cleanupVocabularyStats(db, {
      resolveLegacyPos: async (row) => {
        if (row.id !== legacyId) {
          return null;
        }
        return {
          partOfSpeech: 'verb',
          headword: '知る',
          reading: 'しっている',
          pos1: '動詞',
          pos2: '自立',
          pos3: '',
        };
      },
    });

    const rows = db
      .prepare(
        `SELECT id, headword, word, reading, frequency, first_seen, last_seen
         FROM imm_words
         ORDER BY id ASC`,
      )
      .all() as Array<{
      id: number;
      headword: string;
      word: string;
      reading: string;
      frequency: number;
      first_seen: number;
      last_seen: number;
    }>;
    const occurrences = getWordOccurrences(db, '知る', '知っている', 'しっている', 10);

    assert.deepEqual(result, { scanned: 2, kept: 1, deleted: 1, repaired: 1 });
    assert.deepEqual(rows, [
      {
        id: duplicateId,
        headword: '知る',
        word: '知っている',
        reading: 'しっている',
        frequency: 7,
        first_seen: 1_000,
        last_seen: 3_000,
      },
    ]);
    assert.deepEqual(occurrences, [
      {
        animeId: null,
        animeTitle: null,
        videoId,
        videoTitle: 'Cleanup Merge',
        sessionId,
        lineIndex: 1,
        segmentStartMs: 0,
        segmentEndMs: 1000,
        text: '知っている',
        occurrenceCount: 3,
      },
    ]);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getKanjiStats returns rows ordered by frequency descending', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    // Insert kanji with highest-frequency entry inserted after another character
    stmts.kanjiUpsertStmt.run('月', 1_500, 1_500);
    stmts.kanjiUpsertStmt.run('日', 1_000, 2_000);
    stmts.kanjiUpsertStmt.run('日', 1_000, 3_000);

    const rows = getKanjiStats(db, 10);

    assert.equal(rows.length, 2);
    assert.equal(rows[0]?.kanji, '日');
    assert.equal(rows[1]?.kanji, '月');

    assert.ok(rows.length >= 2);
    const nichiRow = rows.find((r) => r.kanji === '日');
    const tsukiRow = rows.find((r) => r.kanji === '月');
    assert.ok(nichiRow, 'expected 日 row');
    assert.ok(tsukiRow, 'expected 月 row');
    assert.equal(nichiRow.kanji, '日');
    assert.equal(nichiRow.frequency, 2);
    assert.equal(typeof nichiRow.firstSeen, 'number');
    assert.equal(typeof nichiRow.lastSeen, 'number');
    // Higher frequency should come first
    const nichiIdx = rows.indexOf(nichiRow);
    const tsukiIdx = rows.indexOf(tsukiRow);
    assert.ok(nichiIdx < tsukiIdx, 'higher frequency kanji should appear first');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getKanjiStats returns empty array when no kanji exist', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const rows = getKanjiStats(db, 10);
    assert.deepEqual(rows, []);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getSessionEvents returns events ordered by ts_ms ascending', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/events-test.mkv', {
      canonicalTitle: 'Events Test',
      sourcePath: '/tmp/events-test.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const startedAtMs = 5_000_000;
    const { sessionId } = startSessionRecord(db, videoId, startedAtMs);

    // Insert two events at different timestamps
    stmts.eventInsertStmt.run(
      sessionId,
      startedAtMs + 2_000,
      EVENT_SUBTITLE_LINE,
      1,
      0,
      800,
      2,
      0,
      '{"line":"second"}',
      startedAtMs + 2_000,
      startedAtMs + 2_000,
    );
    stmts.eventInsertStmt.run(
      sessionId,
      startedAtMs + 1_000,
      EVENT_SUBTITLE_LINE,
      0,
      0,
      600,
      3,
      0,
      '{"line":"first"}',
      startedAtMs + 1_000,
      startedAtMs + 1_000,
    );

    const events = getSessionEvents(db, sessionId, 50);

    assert.equal(events.length, 2);
    // Should be ordered ASC by ts_ms
    assert.equal(events[0]!.tsMs, startedAtMs + 1_000);
    assert.equal(events[1]!.tsMs, startedAtMs + 2_000);
    assert.equal(events[0]!.eventType, EVENT_SUBTITLE_LINE);
    assert.equal(events[0]!.payload, '{"line":"first"}');
    assert.equal(events[1]!.payload, '{"line":"second"}');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getSessionEvents returns empty array for session with no events', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/events-empty.mkv', {
      canonicalTitle: 'Events Empty',
      sourcePath: '/tmp/events-empty.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const { sessionId } = startSessionRecord(db, videoId, 6_000_000);

    const events = getSessionEvents(db, sessionId, 50);
    assert.deepEqual(events, []);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getSessionEvents filters events to the requested session id', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const decoyVideoId = getOrCreateVideoRecord(db, 'local:/tmp/events-filter-decoy.mkv', {
      canonicalTitle: 'Events Filter Decoy',
      sourcePath: '/tmp/events-filter-decoy.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const targetVideoId = getOrCreateVideoRecord(db, 'local:/tmp/events-filter-target.mkv', {
      canonicalTitle: 'Events Filter Target',
      sourcePath: '/tmp/events-filter-target.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const decoySession = startSessionRecord(db, decoyVideoId, 8_000_000);
    const targetSession = startSessionRecord(db, targetVideoId, 8_100_000);

    // Decoy session event
    stmts.eventInsertStmt.run(
      decoySession.sessionId,
      8_100_000 + 1,
      EVENT_SUBTITLE_LINE,
      1,
      0,
      500,
      1,
      0,
      '{"line":"decoy"}',
      8_100_000 + 1,
      8_100_000 + 1,
    );

    // Target session event
    stmts.eventInsertStmt.run(
      targetSession.sessionId,
      8_100_000 + 2,
      EVENT_SUBTITLE_LINE,
      2,
      0,
      600,
      1,
      0,
      '{"line":"target"}',
      8_100_000 + 2,
      8_100_000 + 2,
    );

    const events = getSessionEvents(db, targetSession.sessionId, 50);

    assert.equal(events.length, 1);
    assert.equal(events[0]?.payload, '{"line":"target"}');
    assert.equal(events[0]?.eventType, EVENT_SUBTITLE_LINE);
    assert.equal(events[0]?.tsMs, 8100002);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getSessionEvents respects limit parameter', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/events-limit.mkv', {
      canonicalTitle: 'Events Limit Test',
      sourcePath: '/tmp/events-limit.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const startedAtMs = 7_000_000;
    const { sessionId } = startSessionRecord(db, videoId, startedAtMs);

    // Insert 5 events
    for (let i = 0; i < 5; i += 1) {
      stmts.eventInsertStmt.run(
        sessionId,
        startedAtMs + i * 1_000,
        EVENT_SUBTITLE_LINE,
        i,
        0,
        500,
        1,
        0,
        null,
        startedAtMs + i * 1_000,
        startedAtMs + i * 1_000,
      );
    }

    const limited = getSessionEvents(db, sessionId, 3);
    assert.equal(limited.length, 3);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('anime-level queries group by anime_id and preserve episode-level rows', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const lwaAnimeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Little Witch Academia',
      canonicalTitle: 'Little Witch Academia',
      anilistId: 33_435,
      titleRomaji: 'Little Witch Academia',
      titleEnglish: 'Little Witch Academia',
      titleNative: 'リトルウィッチアカデミア',
      metadataJson: '{"source":"anilist"}',
    });
    const frierenAnimeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Frieren',
      canonicalTitle: 'Frieren',
      anilistId: 52_921,
      titleRomaji: 'Sousou no Frieren',
      titleEnglish: 'Frieren: Beyond Journey\'s End',
      titleNative: '葬送のフリーレン',
      metadataJson: '{"source":"anilist"}',
    });

    const lwaEpisode5 = getOrCreateVideoRecord(db, 'local:/tmp/lwa-s02e05.mkv', {
      canonicalTitle: 'Episode 5',
      sourcePath: '/tmp/Little Witch Academia S02E05.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const lwaEpisode6 = getOrCreateVideoRecord(db, 'local:/tmp/lwa-s02e06.mkv', {
      canonicalTitle: 'Episode 6',
      sourcePath: '/tmp/Little Witch Academia S02E06.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const frierenEpisode3 = getOrCreateVideoRecord(db, 'local:/tmp/frieren-03.mkv', {
      canonicalTitle: 'Episode 3',
      sourcePath: '/tmp/[SubsPlease] Frieren - 03 - Departure.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    linkVideoToAnimeRecord(db, lwaEpisode5, {
      animeId: lwaAnimeId,
      parsedBasename: 'Little Witch Academia S02E05.mkv',
      parsedTitle: 'Little Witch Academia',
      parsedSeason: 2,
      parsedEpisode: 5,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":5}',
    });
    linkVideoToAnimeRecord(db, lwaEpisode6, {
      animeId: lwaAnimeId,
      parsedBasename: 'Little Witch Academia S02E06.mkv',
      parsedTitle: 'Little Witch Academia',
      parsedSeason: 2,
      parsedEpisode: 6,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":6}',
    });
    linkVideoToAnimeRecord(db, frierenEpisode3, {
      animeId: frierenAnimeId,
      parsedBasename: '[SubsPlease] Frieren - 03 - Departure.mkv',
      parsedTitle: 'Frieren',
      parsedSeason: 1,
      parsedEpisode: 3,
      parserSource: 'fallback',
      parserConfidence: 0.6,
      parseMetadataJson: '{"episode":3}',
    });

    const sessionA = startSessionRecord(db, lwaEpisode5, 1_000_000);
    const sessionB = startSessionRecord(db, lwaEpisode5, 1_010_000);
    const sessionC = startSessionRecord(db, lwaEpisode6, 1_020_000);
    const sessionD = startSessionRecord(db, frierenEpisode3, 1_030_000);

    stmts.telemetryInsertStmt.run(
      sessionA.sessionId,
      1_001_000,
      4_000,
      3_000,
      10,
      25,
      25,
      1,
      3,
      2,
      0,
      0,
      0,
      0,
      0,
      1_001_000,
      1_001_000,
    );
    stmts.telemetryInsertStmt.run(
      sessionB.sessionId,
      1_011_000,
      5_000,
      4_000,
      11,
      27,
      27,
      2,
      4,
      2,
      0,
      0,
      0,
      0,
      0,
      1_011_000,
      1_011_000,
    );
    stmts.telemetryInsertStmt.run(
      sessionC.sessionId,
      1_021_000,
      6_000,
      5_000,
      12,
      28,
      28,
      3,
      5,
      4,
      0,
      0,
      0,
      0,
      0,
      1_021_000,
      1_021_000,
    );
    stmts.telemetryInsertStmt.run(
      sessionD.sessionId,
      1_031_000,
      4_000,
      3_500,
      8,
      20,
      20,
      1,
      2,
      1,
      0,
      0,
      0,
      0,
      0,
      1_031_000,
      1_031_000,
    );

    const animeLibrary = getAnimeLibrary(db);
    assert.equal(animeLibrary.length, 2);
    assert.deepEqual(
      animeLibrary.map((row) => ({
        animeId: row.animeId,
        canonicalTitle: row.canonicalTitle,
        totalSessions: row.totalSessions,
        totalActiveMs: row.totalActiveMs,
        totalCards: row.totalCards,
        episodeCount: row.episodeCount,
      })),
      [
        {
          animeId: lwaAnimeId,
          canonicalTitle: 'Little Witch Academia',
          totalSessions: 3,
          totalActiveMs: 12_000,
          totalCards: 6,
          episodeCount: 2,
        },
        {
          animeId: frierenAnimeId,
          canonicalTitle: 'Frieren',
          totalSessions: 1,
          totalActiveMs: 3_500,
          totalCards: 1,
          episodeCount: 1,
        },
      ],
    );

    const animeDetail = getAnimeDetail(db, lwaAnimeId);
    assert.ok(animeDetail);
    assert.equal(animeDetail?.animeId, lwaAnimeId);
    assert.equal(animeDetail?.canonicalTitle, 'Little Witch Academia');
    assert.equal(animeDetail?.anilistId, 33_435);
    assert.equal(animeDetail?.totalSessions, 3);
    assert.equal(animeDetail?.totalActiveMs, 12_000);
    assert.equal(animeDetail?.totalCards, 6);
    assert.equal(animeDetail?.totalWordsSeen, 80);
    assert.equal(animeDetail?.totalLinesSeen, 33);
    assert.equal(animeDetail?.totalLookupCount, 12);
    assert.equal(animeDetail?.totalLookupHits, 8);
    assert.equal(animeDetail?.episodeCount, 2);

    const episodes = getAnimeEpisodes(db, lwaAnimeId);
    assert.deepEqual(
      episodes.map((row) => ({
        videoId: row.videoId,
        season: row.season,
        episode: row.episode,
        totalSessions: row.totalSessions,
        totalActiveMs: row.totalActiveMs,
        totalCards: row.totalCards,
      })),
      [
        {
          videoId: lwaEpisode5,
          season: 2,
          episode: 5,
          totalSessions: 2,
          totalActiveMs: 7_000,
          totalCards: 3,
        },
        {
          videoId: lwaEpisode6,
          season: 2,
          episode: 6,
          totalSessions: 1,
          totalActiveMs: 5_000,
          totalCards: 3,
        },
      ],
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getWordOccurrences maps a normalized word back to anime, video, and subtitle line context', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Little Witch Academia',
      canonicalTitle: 'Little Witch Academia',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: '{"source":"test"}',
    });
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/lwa-s02e04.mkv', {
      canonicalTitle: 'Episode 4',
      sourcePath: '/tmp/Little Witch Academia S02E04.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: 'Little Witch Academia S02E04.mkv',
      parsedTitle: 'Little Witch Academia',
      parsedSeason: 2,
      parsedEpisode: 4,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":4}',
    });
    const { sessionId } = startSessionRecord(db, videoId, 1_000_000);
    const wordResult = db
      .prepare(
        `INSERT INTO imm_words (
          headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run('猫', '猫', 'ねこ', 'noun', '名詞', '一般', '', 1_000, 1_500, 4);
    const lineResult = db
      .prepare(
        `INSERT INTO imm_subtitle_lines (
          session_id, event_id, video_id, anime_id, line_index, segment_start_ms, segment_end_ms, text, CREATED_DATE, LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run(sessionId, null, videoId, animeId, 1, 0, 1000, '猫 猫 日 日 は', 1_000, 1_000);
    db.prepare(
      `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(Number(lineResult.lastInsertRowid), Number(wordResult.lastInsertRowid), 2);

    const rows = getWordOccurrences(db, '猫', '猫', 'ねこ', 10);

    assert.deepEqual(rows, [
      {
        animeId,
        animeTitle: 'Little Witch Academia',
        videoId,
        videoTitle: 'Episode 4',
        sessionId,
        lineIndex: 1,
        segmentStartMs: 0,
        segmentEndMs: 1000,
        text: '猫 猫 日 日 は',
        occurrenceCount: 2,
      },
    ]);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getKanjiOccurrences maps a kanji back to anime, video, and subtitle line context', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Frieren',
      canonicalTitle: 'Frieren',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: '{"source":"test"}',
    });
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/frieren-03.mkv', {
      canonicalTitle: 'Episode 3',
      sourcePath: '/tmp/[SubsPlease] Frieren - 03 - Departure.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: '[SubsPlease] Frieren - 03 - Departure.mkv',
      parsedTitle: 'Frieren',
      parsedSeason: 1,
      parsedEpisode: 3,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":3}',
    });
    const { sessionId } = startSessionRecord(db, videoId, 2_000_000);
    const kanjiResult = db
      .prepare(
        `INSERT INTO imm_kanji (
          kanji, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?)`,
      )
      .run('日', 2_000, 2_500, 8);
    const lineResult = db
      .prepare(
        `INSERT INTO imm_subtitle_lines (
          session_id, event_id, video_id, anime_id, line_index, segment_start_ms, segment_end_ms, text, CREATED_DATE, LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run(sessionId, null, videoId, animeId, 3, 5000, 6500, '今日は日曜', 2_000, 2_000);
    db.prepare(
      `INSERT INTO imm_kanji_line_occurrences (line_id, kanji_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(Number(lineResult.lastInsertRowid), Number(kanjiResult.lastInsertRowid), 2);

    const rows = getKanjiOccurrences(db, '日', 10);

    assert.deepEqual(rows, [
      {
        animeId,
        animeTitle: 'Frieren',
        videoId,
        videoTitle: 'Episode 3',
        sessionId,
        lineIndex: 3,
        segmentStartMs: 5000,
        segmentEndMs: 6500,
        text: '今日は日曜',
        occurrenceCount: 2,
      },
    ]);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});
