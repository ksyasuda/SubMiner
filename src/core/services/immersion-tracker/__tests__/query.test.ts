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
  getAnimeDailyRollups,
  cleanupVocabularyStats,
  deleteSession,
  getDailyRollups,
  getTrendsDashboard,
  getQueryHints,
  getMonthlyRollups,
  getAnimeDetail,
  getAnimeEpisodes,
  getAnimeCoverArt,
  getAnimeLibrary,
  getCoverArt,
  getMediaDetail,
  getMediaLibrary,
  getKanjiOccurrences,
  getSessionSummaries,
  getVocabularyStats,
  getKanjiStats,
  getSessionEvents,
  getSessionTimeline,
  getSessionWordsByLine,
  getWordOccurrences,
  upsertCoverArt,
} from '../query.js';
import {
  SOURCE_TYPE_LOCAL,
  SOURCE_TYPE_REMOTE,
  EVENT_CARD_MINED,
  EVENT_SUBTITLE_LINE,
  EVENT_YOMITAN_LOOKUP,
} from '../types.js';

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
    assert.equal(row.tokensSeen, 10);
    assert.equal(row.lookupCount, 2);
    assert.equal(row.lookupHits, 1);
    assert.equal(row.yomitanLookupCount, 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getAnimeEpisodes prefers the latest session media position when the latest session is still active', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/active-progress-episode.mkv', {
      canonicalTitle: 'Active Progress Episode',
      sourcePath: '/tmp/active-progress-episode.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Active Progress Anime',
      canonicalTitle: 'Active Progress Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: 'active-progress-episode.mkv',
      parsedTitle: 'Active Progress Anime',
      parsedSeason: 1,
      parsedEpisode: 2,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":2}',
    });

    const endedSessionId = startSessionRecord(db, videoId, 1_000_000).sessionId;
    const activeSessionId = startSessionRecord(db, videoId, 1_010_000).sessionId;
    db.prepare(
      `
      UPDATE imm_sessions
      SET
        ended_at_ms = ?,
        status = 2,
        ended_media_ms = ?,
        active_watched_ms = ?,
        LAST_UPDATE_DATE = ?
      WHERE session_id = ?
      `,
    ).run(1_005_000, 6_000, 3_000, 1_005_000, endedSessionId);
    db.prepare(
      `
      UPDATE imm_sessions
      SET
        ended_media_ms = ?,
        active_watched_ms = ?,
        LAST_UPDATE_DATE = ?
      WHERE session_id = ?
      `,
    ).run(9_000, 4_000, 1_012_000, activeSessionId);

    const [episode] = getAnimeEpisodes(db, animeId);
    assert.ok(episode);
    assert.equal(episode?.endedMediaMs, 9_000);
    assert.equal(episode?.totalSessions, 2);
    assert.equal(episode?.totalActiveMs, 7_000);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getAnimeEpisodes falls back to the latest subtitle segment end when session progress checkpoints are missing', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/subtitle-progress-fallback.mkv', {
      canonicalTitle: 'Subtitle Progress Fallback',
      sourcePath: '/tmp/subtitle-progress-fallback.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Subtitle Progress Fallback Anime',
      canonicalTitle: 'Subtitle Progress Fallback Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: 'subtitle-progress-fallback.mkv',
      parsedTitle: 'Subtitle Progress Fallback Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":1}',
    });
    db.prepare('UPDATE imm_videos SET duration_ms = ? WHERE video_id = ?').run(24_000, videoId);

    const startedAtMs = 1_100_000;
    const sessionId = startSessionRecord(db, videoId, startedAtMs).sessionId;
    db.prepare(
      `
      UPDATE imm_sessions
      SET
        ended_at_ms = ?,
        status = 2,
        active_watched_ms = ?,
        LAST_UPDATE_DATE = ?
      WHERE session_id = ?
      `,
    ).run(startedAtMs + 10_000, 10_000, startedAtMs + 10_000, sessionId);
    stmts.eventInsertStmt.run(
      sessionId,
      startedAtMs + 9_000,
      EVENT_SUBTITLE_LINE,
      1,
      18_000,
      21_000,
      5,
      0,
      '{"line":"progress fallback"}',
      startedAtMs + 9_000,
      startedAtMs + 9_000,
    );

    const [episode] = getAnimeEpisodes(db, animeId);
    assert.ok(episode);
    assert.equal(episode?.endedMediaMs, 21_000);
    assert.equal(episode?.totalSessions, 1);
    assert.equal(episode?.totalActiveMs, 10_000);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getSessionTimeline returns the full session when no limit is provided', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/full-timeline-test.mkv', {
      canonicalTitle: 'Full Timeline Test',
      sourcePath: '/tmp/full-timeline-test.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const startedAtMs = 2_000_000;
    const { sessionId } = startSessionRecord(db, videoId, startedAtMs);

    for (let sample = 0; sample < 205; sample += 1) {
      const sampleMs = startedAtMs + sample * 500;
      stmts.telemetryInsertStmt.run(
        sessionId,
        sampleMs,
        sample * 500,
        sample * 450,
        sample,
        sample * 4,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        sampleMs,
        sampleMs,
      );
    }

    const rows = getSessionTimeline(db, sessionId);

    assert.equal(rows.length, 205);
    assert.equal(rows[0]?.linesSeen, 204);
    assert.equal(rows.at(-1)?.linesSeen, 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getDailyRollups limits by distinct days (not rows)', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const insert = db.prepare(
      `
      INSERT INTO imm_daily_rollups (
        rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
        total_tokens_seen, total_cards
      ) VALUES (?, ?, ?, ?, ?, ?, ?)
    `,
    );

    insert.run(10, 1, 1, 1, 0, 0, 2);
    insert.run(10, 2, 1, 1, 0, 0, 3);
    insert.run(9, 1, 1, 1, 0, 0, 1);
    insert.run(8, 1, 1, 1, 0, 0, 1);

    const rows = getDailyRollups(db, 2);
    assert.equal(rows.length, 3);
    assert.ok(rows.every((r) => r.rollupDayOrMonth === 10 || r.rollupDayOrMonth === 9));
    assert.ok(rows.some((r) => r.rollupDayOrMonth === 10 && r.videoId === 1));
    assert.ok(rows.some((r) => r.rollupDayOrMonth === 10 && r.videoId === 2));
    assert.ok(rows.some((r) => r.rollupDayOrMonth === 9 && r.videoId === 1));
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getTrendsDashboard returns chart-ready aggregated series', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/trends-dashboard-test.mkv', {
      canonicalTitle: 'Trend Dashboard Test',
      sourcePath: '/tmp/trends-dashboard-test.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Trend Dashboard Anime',
      canonicalTitle: 'Trend Dashboard Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: 'trends-dashboard-test.mkv',
      parsedTitle: 'Trend Dashboard Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'test',
      parserConfidence: 1,
      parseMetadataJson: null,
    });

    const dayOneStart = new Date(2026, 2, 15, 12, 0, 0, 0).getTime();
    const dayTwoStart = new Date(2026, 2, 16, 18, 0, 0, 0).getTime();

    const sessionOne = startSessionRecord(db, videoId, dayOneStart);
    const sessionTwo = startSessionRecord(db, videoId, dayTwoStart);

    for (const [
      sessionId,
      startedAtMs,
      activeWatchedMs,
      cardsMined,
      tokensSeen,
      yomitanLookupCount,
    ] of [
      [sessionOne.sessionId, dayOneStart, 30 * 60_000, 2, 120, 8],
      [sessionTwo.sessionId, dayTwoStart, 45 * 60_000, 3, 140, 10],
    ] as const) {
      stmts.telemetryInsertStmt.run(
        sessionId,
        startedAtMs + 60_000,
        activeWatchedMs,
        activeWatchedMs,
        10,
        tokensSeen,
        cardsMined,
        0,
        0,
        yomitanLookupCount,
        0,
        0,
        0,
        0,
        startedAtMs + 60_000,
        startedAtMs + 60_000,
      );

      db.prepare(
        `
          UPDATE imm_sessions
          SET
            ended_at_ms = ?,
            total_watched_ms = ?,
            active_watched_ms = ?,
            lines_seen = ?,
            tokens_seen = ?,
            cards_mined = ?,
            yomitan_lookup_count = ?
          WHERE session_id = ?
        `,
      ).run(
        startedAtMs + activeWatchedMs,
        activeWatchedMs,
        activeWatchedMs,
        10,
        tokensSeen,
        cardsMined,
        yomitanLookupCount,
        sessionId,
      );
    }

    db.prepare(
      `
        INSERT INTO imm_daily_rollups (
          rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
          total_tokens_seen, total_cards
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(Math.floor(dayOneStart / 86_400_000), videoId, 1, 30, 10, 120, 2);

    db.prepare(
      `
        INSERT INTO imm_daily_rollups (
          rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
          total_tokens_seen, total_cards
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(Math.floor(dayTwoStart / 86_400_000), videoId, 1, 45, 10, 140, 3);

    db.prepare(
      `
        INSERT INTO imm_words (
          headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(
      '勉強',
      '勉強',
      'べんきょう',
      'noun',
      '名詞',
      null,
      null,
      Math.floor(dayOneStart / 1000),
      Math.floor(dayTwoStart / 1000),
    );

    const dashboard = getTrendsDashboard(db, 'all', 'day');

    assert.equal(dashboard.activity.watchTime.length, 2);
    assert.equal(dashboard.activity.watchTime[0]?.value, 30);
    assert.equal(dashboard.progress.watchTime[1]?.value, 75);
    assert.equal(dashboard.progress.lookups[1]?.value, 18);
    assert.equal(dashboard.ratios.lookupsPerHundred[0]?.value, +((8 / 120) * 100).toFixed(1));
    assert.equal(dashboard.animePerDay.watchTime[0]?.animeTitle, 'Trend Dashboard Anime');
    assert.equal(dashboard.animeCumulative.watchTime[1]?.value, 75);
    assert.equal(
      dashboard.patterns.watchTimeByDayOfWeek.reduce((sum, point) => sum + point.value, 0),
      75,
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getQueryHints reads all-time totals from lifetime summary', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    db.prepare(
      `
      UPDATE imm_lifetime_global
      SET
        total_sessions = ?,
        total_active_ms = ?,
        total_cards = ?,
        active_days = ?,
        episodes_completed = ?,
        anime_completed = ?
      WHERE global_id = 1
      `,
    ).run(4, 90_000, 2, 9, 11, 22);

    const insert = db.prepare(
      `
      INSERT INTO imm_daily_rollups (
        rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
        total_tokens_seen, total_cards
      ) VALUES (?, ?, ?, ?, ?, ?, ?)
    `,
    );

    insert.run(10, 1, 1, 12, 0, 0, 2);
    insert.run(10, 2, 1, 11, 0, 0, 3);
    insert.run(9, 1, 1, 10, 0, 0, 1);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/query-hints.mkv', {
      canonicalTitle: 'Query Hints Episode',
      sourcePath: '/tmp/query-hints.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const { sessionId } = startSessionRecord(db, videoId, 1_000_000);
    db.prepare(
      `
      UPDATE imm_sessions
      SET
        ended_at_ms = ?,
        status = 2,
        tokens_seen = ?,
        yomitan_lookup_count = ?,
        lookup_count = ?,
        lookup_hits = ?,
        LAST_UPDATE_DATE = ?
      WHERE session_id = ?
      `,
    ).run(1_060_000, 120, 8, 11, 7, 1_060_000, sessionId);

    const hints = getQueryHints(db);
    assert.equal(hints.totalSessions, 4);
    assert.equal(hints.totalCards, 2);
    assert.equal(hints.totalActiveMin, 1);
    assert.equal(hints.activeDays, 9);
    assert.equal(hints.totalEpisodesWatched, 11);
    assert.equal(hints.totalAnimeCompleted, 22);
    assert.equal(hints.totalTokensSeen, 120);
    assert.equal(hints.totalYomitanLookupCount, 8);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getQueryHints counts new words by distinct headword first-seen time', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const now = new Date();
    const todayStartSec =
      new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime() / 1000;
    const oneHourAgo = todayStartSec + 3_600;
    const twoDaysAgo = todayStartSec - 2 * 86_400;

    db.prepare(
      `
      INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run('知る', '知った', 'しった', 'verb', '動詞', '', '', oneHourAgo, oneHourAgo, 1);
    db.prepare(
      `
      INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run('知る', '知っている', 'しっている', 'verb', '動詞', '', '', oneHourAgo, oneHourAgo, 1);
    db.prepare(
      `
      INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run('猫', '猫', 'ねこ', 'noun', '名詞', '', '', twoDaysAgo, twoDaysAgo, 1);

    const hints = getQueryHints(db);
    assert.equal(hints.newWordsToday, 1);
    assert.equal(hints.newWordsThisWeek, 2);
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
    assert.equal(row.tokensSeen, 0);
    assert.equal(row.lookupCount, 0);
    assert.equal(row.lookupHits, 0);
    assert.equal(row.yomitanLookupCount, 0);
    assert.equal(row.cardsMined, 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getSessionSummaries uses denormalized session metrics for ended sessions without telemetry', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/ended-session-no-telemetry.mkv', {
      canonicalTitle: 'Ended Session',
      sourcePath: '/tmp/ended-session-no-telemetry.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const startedAtMs = 4_000_000;
    const endedAtMs = startedAtMs + 8_000;
    const { sessionId } = startSessionRecord(db, videoId, startedAtMs);
    db.prepare(
      `
      UPDATE imm_sessions
      SET
        ended_at_ms = ?,
        status = 2,
        total_watched_ms = ?,
        active_watched_ms = ?,
        lines_seen = ?,
        tokens_seen = ?,
        cards_mined = ?,
        lookup_count = ?,
        lookup_hits = ?,
        LAST_UPDATE_DATE = ?
      WHERE session_id = ?
      `,
    ).run(endedAtMs, 8_000, 7_000, 12, 34, 5, 9, 6, endedAtMs, sessionId);

    const rows = getSessionSummaries(db, 10);
    const row = rows.find((r) => r.sessionId === sessionId);
    assert.ok(row);
    assert.equal(row.totalWatchedMs, 8_000);
    assert.equal(row.activeWatchedMs, 7_000);
    assert.equal(row.linesSeen, 12);
    assert.equal(row.tokensSeen, 34);
    assert.equal(row.cardsMined, 5);
    assert.equal(row.lookupCount, 9);
    assert.equal(row.lookupHits, 6);
    assert.equal(row.yomitanLookupCount, 0);
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
    assert.deepEqual(repairedRows, [
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
    ]);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getDailyRollups returns all rows for the most recent rollup days', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  try {
    ensureSchema(db);
    const insertRollup = db.prepare(
      `
      INSERT INTO imm_daily_rollups (
        rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
        total_tokens_seen, total_cards, cards_per_hour, tokens_per_min, lookup_hit_rate
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `,
    );
    insertRollup.run(3_000, 1, 1, 10, 20, 40, 2, 0.1, 0.2, 0.3);
    insertRollup.run(3_000, 2, 2, 10, 20, 40, 3, 0.1, 0.2, 0.3);
    insertRollup.run(2_999, 3, 1, 5, 10, 20, 1, 0.1, 0.2, 0.3);
    insertRollup.run(2_998, 4, 1, 5, 10, 20, 1, 0.1, 0.2, 0.3);

    const rows = getDailyRollups(db, 1);
    assert.equal(rows.length, 2);
    assert.equal(rows[0]?.rollupDayOrMonth, 3_000);
    assert.equal(rows[0]?.videoId, 2);
    assert.equal(rows[1]?.rollupDayOrMonth, 3_000);
    assert.equal(rows[1]?.videoId, 1);

    const twoRows = getDailyRollups(db, 2);
    assert.equal(twoRows.length, 3);
    assert.equal(twoRows[0]?.rollupDayOrMonth, 3_000);
    assert.equal(twoRows[1]?.rollupDayOrMonth, 3_000);
    assert.equal(twoRows[2]?.rollupDayOrMonth, 2_999);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getMonthlyRollups returns all rows for the most recent rollup months', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  try {
    ensureSchema(db);
    const insertRollup = db.prepare(
      `
      INSERT INTO imm_monthly_rollups (
        rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
        total_tokens_seen, total_cards, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `,
    );
    const nowMs = Date.now();
    insertRollup.run(202602, 1, 1, 10, 20, 40, 5, nowMs, nowMs);
    insertRollup.run(202602, 2, 1, 10, 20, 40, 6, nowMs, nowMs);
    insertRollup.run(202601, 3, 1, 5, 10, 20, 2, nowMs, nowMs);
    insertRollup.run(202600, 4, 1, 5, 10, 20, 2, nowMs, nowMs);

    const rows = getMonthlyRollups(db, 1);
    assert.equal(rows.length, 2);
    assert.equal(rows[0]?.rollupDayOrMonth, 202602);
    assert.equal(rows[0]?.videoId, 2);
    assert.equal(rows[1]?.rollupDayOrMonth, 202602);
    assert.equal(rows[1]?.videoId, 1);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getAnimeDailyRollups returns all rows for the most recent rollup days', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  try {
    ensureSchema(db);
    const insertRollup = db.prepare(
      `
      INSERT INTO imm_daily_rollups (
        rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
        total_tokens_seen, total_cards, cards_per_hour, tokens_per_min, lookup_hit_rate
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `,
    );
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Test Anime',
      canonicalTitle: 'Test Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    const video1 = getOrCreateVideoRecord(db, 'local:/tmp/anime-ep1.mkv', {
      canonicalTitle: 'Episode 1',
      sourcePath: '/tmp/anime-ep1.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const video2 = getOrCreateVideoRecord(db, 'local:/tmp/anime-ep2.mkv', {
      canonicalTitle: 'Episode 2',
      sourcePath: '/tmp/anime-ep2.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    db.prepare('UPDATE imm_videos SET anime_id = ? WHERE video_id IN (?, ?)').run(
      animeId,
      video1,
      video2,
    );

    insertRollup.run(4_000, video1, 1, 10, 20, 40, 2, 0.1, 0.2, 0.3);
    insertRollup.run(4_000, video2, 1, 10, 20, 40, 2, 0.1, 0.2, 0.3);
    insertRollup.run(3_999, video1, 1, 10, 20, 40, 2, 0.1, 0.2, 0.3);

    const rows = getAnimeDailyRollups(db, animeId, 1);
    assert.equal(rows.length, 2);
    assert.equal(rows[0]?.rollupDayOrMonth, 4_000);
    assert.equal(rows[0]?.videoId, video2);
    assert.equal(rows[1]?.rollupDayOrMonth, 4_000);
    assert.equal(rows[1]?.videoId, video1);
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
        sourcePath: '/tmp/cleanup-merge.mkv',
        secondaryText: null,
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

test('getSessionEvents filters by event type before applying limit', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/events-type-filter.mkv', {
      canonicalTitle: 'Events Type Filter',
      sourcePath: '/tmp/events-type-filter.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const startedAtMs = 7_500_000;
    const { sessionId } = startSessionRecord(db, videoId, startedAtMs);

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
        `{"line":"subtitle-${i}"}`,
        startedAtMs + i * 1_000,
        startedAtMs + i * 1_000,
      );
    }

    stmts.eventInsertStmt.run(
      sessionId,
      startedAtMs + 10_000,
      EVENT_CARD_MINED,
      null,
      null,
      null,
      0,
      1,
      '{"cardsMined":1}',
      startedAtMs + 10_000,
      startedAtMs + 10_000,
    );

    stmts.eventInsertStmt.run(
      sessionId,
      startedAtMs + 11_000,
      EVENT_YOMITAN_LOOKUP,
      null,
      null,
      null,
      0,
      0,
      null,
      startedAtMs + 11_000,
      startedAtMs + 11_000,
    );

    const filtered = getSessionEvents(db, sessionId, 1, [EVENT_CARD_MINED]);
    assert.equal(filtered.length, 1);
    assert.equal(filtered[0]?.eventType, EVENT_CARD_MINED);
    assert.equal(filtered[0]?.payload, '{"cardsMined":1}');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('getSessionWordsByLine joins word occurrences through imm_words.id', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);
    const startedAtMs = Date.UTC(2025, 0, 1, 12, 0, 0);
    const videoId = getOrCreateVideoRecord(db, '/tmp/session-words-by-line.mkv', {
      canonicalTitle: 'Episode',
      sourcePath: '/tmp/session-words-by-line.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const { sessionId } = startSessionRecord(db, videoId, startedAtMs);
    const lineId = Number(
      db
        .prepare(
          `INSERT INTO imm_subtitle_lines (
            session_id, event_id, video_id, anime_id, line_index, segment_start_ms, segment_end_ms, text, CREATED_DATE, LAST_UPDATE_DATE
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        )
        .run(sessionId, null, videoId, null, 0, 0, 1000, '猫を見た', startedAtMs, startedAtMs)
        .lastInsertRowid,
    );
    const wordId = Number(
      db
        .prepare(
          `INSERT INTO imm_words (
            headword, word, reading, pos1, pos2, pos3, part_of_speech, first_seen, last_seen, frequency
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        )
        .run('猫', '猫', 'ねこ', null, null, null, null, startedAtMs, startedAtMs, 1)
        .lastInsertRowid,
    );

    db.prepare(
      `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(lineId, wordId, 1);

    assert.deepEqual(getSessionWordsByLine(db, sessionId), [
      { lineIndex: 0, headword: '猫', occurrenceCount: 1 },
    ]);
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
      titleEnglish: "Frieren: Beyond Journey's End",
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

    const now = Date.now();
    db.prepare(
      `
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
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `,
    ).run(lwaAnimeId, 3, 12_000, 6, 33, 80, 2, 1, 1_000_000, 1_021_000, now, now);
    db.prepare(
      `
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
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `,
    ).run(frierenAnimeId, 1, 3_500, 1, 8, 20, 1, 1, 1_030_000, 1_030_000, now, now);

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
    assert.equal(animeDetail?.totalTokensSeen, 80);
    assert.equal(animeDetail?.totalLinesSeen, 33);
    assert.equal(animeDetail?.totalLookupCount, 12);
    assert.equal(animeDetail?.totalLookupHits, 8);
    assert.equal(animeDetail?.totalYomitanLookupCount, 0);
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
        totalTokensSeen: row.totalTokensSeen,
        totalYomitanLookupCount: row.totalYomitanLookupCount,
      })),
      [
        {
          videoId: lwaEpisode5,
          season: 2,
          episode: 5,
          totalSessions: 2,
          totalActiveMs: 7_000,
          totalCards: 3,
          totalTokensSeen: 52,
          totalYomitanLookupCount: 0,
        },
        {
          videoId: lwaEpisode6,
          season: 2,
          episode: 6,
          totalSessions: 1,
          totalActiveMs: 5_000,
          totalCards: 3,
          totalTokensSeen: 28,
          totalYomitanLookupCount: 0,
        },
      ],
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('anime library and detail still return lifetime rows without retained sessions', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'No Session Anime',
      canonicalTitle: 'No Session Anime',
      anilistId: 111_111,
      titleRomaji: 'No Session Anime',
      titleEnglish: 'No Session Anime',
      titleNative: 'No Session Anime',
      metadataJson: null,
    });
    const ep1 = getOrCreateVideoRecord(db, 'local:/tmp/no-session-ep1.mkv', {
      canonicalTitle: 'Episode 1',
      sourcePath: '/tmp/no-session-ep1.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const ep2 = getOrCreateVideoRecord(db, 'local:/tmp/no-session-ep2.mkv', {
      canonicalTitle: 'Episode 2',
      sourcePath: '/tmp/no-session-ep2.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    linkVideoToAnimeRecord(db, ep1, {
      animeId,
      parsedBasename: 'Episode 1',
      parsedTitle: 'No Session Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":1}',
    });
    linkVideoToAnimeRecord(db, ep2, {
      animeId,
      parsedBasename: 'Episode 2',
      parsedTitle: 'No Session Anime',
      parsedSeason: 1,
      parsedEpisode: 2,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":2}',
    });

    const now = Date.now();
    db.prepare(
      `
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
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(animeId, 12, 4_500, 9, 80, 200, 2, 2, 1_000_000, now, now, now);

    const library = getAnimeLibrary(db);
    assert.equal(library.length, 1);
    assert.equal(library[0]?.animeId, animeId);
    assert.equal(library[0]?.canonicalTitle, 'No Session Anime');
    assert.equal(library[0]?.totalSessions, 12);
    assert.equal(library[0]?.totalActiveMs, 4_500);
    assert.equal(library[0]?.totalCards, 9);
    assert.equal(library[0]?.episodeCount, 2);

    const detail = getAnimeDetail(db, animeId);
    assert.ok(detail);
    assert.equal(detail?.animeId, animeId);
    assert.equal(detail?.canonicalTitle, 'No Session Anime');
    assert.equal(detail?.totalSessions, 12);
    assert.equal(detail?.totalActiveMs, 4_500);
    assert.equal(detail?.totalCards, 9);
    assert.equal(detail?.totalTokensSeen, 200);
    assert.equal(detail?.totalLinesSeen, 80);
    assert.equal(detail?.episodeCount, 2);
    assert.equal(detail?.totalLookupCount, 0);
    assert.equal(detail?.totalLookupHits, 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('media library and detail queries read lifetime totals', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const mediaOne = getOrCreateVideoRecord(db, 'local:/tmp/media-one.mkv', {
      canonicalTitle: 'Media One',
      sourcePath: '/tmp/media-one.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const mediaTwo = getOrCreateVideoRecord(db, 'local:/tmp/media-two.mkv', {
      canonicalTitle: 'Media Two',
      sourcePath: '/tmp/media-two.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const insertLifetime = db.prepare(
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
    );
    const now = Date.now();
    const older = now - 10_000;
    insertLifetime.run(mediaOne, 3, 12_000, 4, 10, 180, 1, 1_000, now, now, now);
    insertLifetime.run(mediaTwo, 1, 2_000, 2, 4, 40, 0, 900, older, now, now);

    const library = getMediaLibrary(db);
    assert.equal(library.length, 2);
    assert.deepEqual(
      library.map((row) => ({
        videoId: row.videoId,
        canonicalTitle: row.canonicalTitle,
        totalSessions: row.totalSessions,
        totalActiveMs: row.totalActiveMs,
        totalCards: row.totalCards,
        totalTokensSeen: row.totalTokensSeen,
        lastWatchedMs: row.lastWatchedMs,
        hasCoverArt: row.hasCoverArt,
      })),
      [
        {
          videoId: mediaOne,
          canonicalTitle: 'Media One',
          totalSessions: 3,
          totalActiveMs: 12_000,
          totalCards: 4,
          totalTokensSeen: 180,
          lastWatchedMs: now,
          hasCoverArt: 0,
        },
        {
          videoId: mediaTwo,
          canonicalTitle: 'Media Two',
          totalSessions: 1,
          totalActiveMs: 2_000,
          totalCards: 2,
          totalTokensSeen: 40,
          lastWatchedMs: older,
          hasCoverArt: 0,
        },
      ],
    );

    const detail = getMediaDetail(db, mediaOne);
    assert.ok(detail);
    assert.equal(detail.totalSessions, 3);
    assert.equal(detail.totalActiveMs, 12_000);
    assert.equal(detail.totalCards, 4);
    assert.equal(detail.totalTokensSeen, 180);
    assert.equal(detail.totalLinesSeen, 10);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('media library and detail queries include joined youtube metadata when present', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const mediaOne = getOrCreateVideoRecord(db, 'yt:https://www.youtube.com/watch?v=abc123', {
      canonicalTitle: 'Local Fallback Title',
      sourcePath: null,
      sourceUrl: 'https://www.youtube.com/watch?v=abc123',
      sourceType: SOURCE_TYPE_REMOTE,
    });

    db.prepare(
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
    ).run(mediaOne, 2, 6_000, 1, 5, 80, 0, 1_000, 9_000, 9_000, 9_000);

    db.prepare(
      `
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
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(
      mediaOne,
      'abc123',
      'https://www.youtube.com/watch?v=abc123',
      'Tracked Video Title',
      'https://i.ytimg.com/vi/abc123/hqdefault.jpg',
      'UCcreator123',
      'Creator Name',
      'https://www.youtube.com/channel/UCcreator123',
      'https://yt3.googleusercontent.com/channel-avatar=s88',
      '@creator',
      'https://www.youtube.com/@creator',
      'Video description',
      '{"source":"test"}',
      10_000,
      10_000,
      10_000,
    );

    const library = getMediaLibrary(db);
    const detail = getMediaDetail(db, mediaOne);

    assert.equal(library.length, 1);
    assert.equal(library[0]?.youtubeVideoId, 'abc123');
    assert.equal(library[0]?.videoTitle, 'Tracked Video Title');
    assert.equal(library[0]?.channelId, 'UCcreator123');
    assert.equal(library[0]?.channelName, 'Creator Name');
    assert.equal(library[0]?.channelUrl, 'https://www.youtube.com/channel/UCcreator123');
    assert.equal(detail?.youtubeVideoId, 'abc123');
    assert.equal(detail?.videoUrl, 'https://www.youtube.com/watch?v=abc123');
    assert.equal(detail?.videoThumbnailUrl, 'https://i.ytimg.com/vi/abc123/hqdefault.jpg');
    assert.equal(detail?.channelThumbnailUrl, 'https://yt3.googleusercontent.com/channel-avatar=s88');
    assert.equal(detail?.uploaderId, '@creator');
    assert.equal(detail?.uploaderUrl, 'https://www.youtube.com/@creator');
    assert.equal(detail?.description, 'Video description');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('cover art queries reuse a shared blob across duplicate anime art rows', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Shared Blob Anime',
      canonicalTitle: 'Shared Blob Anime',
      anilistId: 42_424,
      titleRomaji: 'Shared Blob Anime',
      titleEnglish: 'Shared Blob Anime',
      titleNative: null,
      metadataJson: null,
    });
    const videoOne = getOrCreateVideoRecord(db, 'local:/tmp/shared-blob-1.mkv', {
      canonicalTitle: 'Shared Blob 1',
      sourcePath: '/tmp/shared-blob-1.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const videoTwo = getOrCreateVideoRecord(db, 'local:/tmp/shared-blob-2.mkv', {
      canonicalTitle: 'Shared Blob 2',
      sourcePath: '/tmp/shared-blob-2.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    linkVideoToAnimeRecord(db, videoOne, {
      animeId,
      parsedBasename: 'Shared Blob 1',
      parsedTitle: 'Shared Blob Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: null,
    });
    linkVideoToAnimeRecord(db, videoTwo, {
      animeId,
      parsedBasename: 'Shared Blob 2',
      parsedTitle: 'Shared Blob Anime',
      parsedSeason: 1,
      parsedEpisode: 2,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: null,
    });

    const now = Date.now();
    db.prepare(
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
      ) VALUES (?, 1, 1000, 0, 0, 0, 0, ?, ?, ?, ?)
      `,
    ).run(videoOne, now, now, now, now);
    db.prepare(
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
      ) VALUES (?, 1, 1000, 0, 0, 0, 0, ?, ?, ?, ?)
      `,
    ).run(videoTwo, now, now - 1, now, now);

    upsertCoverArt(db, videoOne, {
      anilistId: 42_424,
      coverUrl: 'https://images.test/shared.jpg',
      coverBlob: Buffer.from([1, 2, 3, 4]),
      titleRomaji: 'Shared Blob Anime',
      titleEnglish: 'Shared Blob Anime',
      episodesTotal: 12,
    });
    upsertCoverArt(db, videoTwo, {
      anilistId: 42_424,
      coverUrl: 'https://images.test/shared.jpg',
      coverBlob: Buffer.from([9, 9, 9, 9]),
      titleRomaji: 'Shared Blob Anime',
      titleEnglish: 'Shared Blob Anime',
      episodesTotal: 12,
    });

    const artOne = getCoverArt(db, videoOne);
    const artTwo = getCoverArt(db, videoTwo);
    const animeArt = getAnimeCoverArt(db, animeId);
    const library = getMediaLibrary(db);

    assert.equal(artOne?.coverBlob?.length, 4);
    assert.equal(artTwo?.coverBlob?.length, 4);
    assert.deepEqual(artOne?.coverBlob, artTwo?.coverBlob);
    assert.equal(animeArt?.coverBlob?.length, 4);
    assert.deepEqual(
      library.map((row) => ({
        videoId: row.videoId,
        hasCoverArt: row.hasCoverArt,
      })),
      [
        { videoId: videoOne, hasCoverArt: 1 },
        { videoId: videoTwo, hasCoverArt: 1 },
      ],
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('anime/media detail and episode queries use ended-session metrics when telemetry rows are absent', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Session Metrics Anime',
      canonicalTitle: 'Session Metrics Anime',
      anilistId: 999_001,
      titleRomaji: 'Session Metrics Anime',
      titleEnglish: 'Session Metrics Anime',
      titleNative: 'Session Metrics Anime',
      metadataJson: null,
    });
    const episodeOne = getOrCreateVideoRecord(db, 'local:/tmp/session-metrics-ep1.mkv', {
      canonicalTitle: 'Episode 1',
      sourcePath: '/tmp/session-metrics-ep1.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const episodeTwo = getOrCreateVideoRecord(db, 'local:/tmp/session-metrics-ep2.mkv', {
      canonicalTitle: 'Episode 2',
      sourcePath: '/tmp/session-metrics-ep2.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    linkVideoToAnimeRecord(db, episodeOne, {
      animeId,
      parsedBasename: 'session-metrics-ep1.mkv',
      parsedTitle: 'Session Metrics Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":1}',
    });
    linkVideoToAnimeRecord(db, episodeTwo, {
      animeId,
      parsedBasename: 'session-metrics-ep2.mkv',
      parsedTitle: 'Session Metrics Anime',
      parsedSeason: 1,
      parsedEpisode: 2,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: '{"episode":2}',
    });

    const now = Date.now();
    db.prepare(
      `
      INSERT INTO imm_lifetime_anime (
        anime_id, total_sessions, total_active_ms, total_cards, total_lines_seen,
        total_tokens_seen, episodes_started, episodes_completed, first_watched_ms, last_watched_ms,
        CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(animeId, 3, 12_000, 6, 24, 60, 2, 2, 1_000_000, 1_020_000, now, now);
    db.prepare(
      `
      INSERT INTO imm_lifetime_media (
        video_id, total_sessions, total_active_ms, total_cards, total_lines_seen,
        total_tokens_seen, completed, first_watched_ms, last_watched_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(episodeOne, 2, 7_000, 3, 12, 30, 1, 1_000_000, 1_010_000, now, now);

    const s1 = startSessionRecord(db, episodeOne, 1_000_000).sessionId;
    const s2 = startSessionRecord(db, episodeOne, 1_010_000).sessionId;
    const s3 = startSessionRecord(db, episodeTwo, 1_020_000).sessionId;
    const updateSession = db.prepare(
      `
      UPDATE imm_sessions
      SET
        ended_at_ms = ?,
        status = 2,
        ended_media_ms = ?,
        active_watched_ms = ?,
        cards_mined = ?,
        tokens_seen = ?,
        lookup_count = ?,
        lookup_hits = ?,
        LAST_UPDATE_DATE = ?
      WHERE session_id = ?
      `,
    );
    updateSession.run(1_001_000, 2_500, 3_000, 1, 10, 4, 3, now, s1);
    updateSession.run(1_011_000, 6_000, 4_000, 2, 20, 5, 4, now, s2);
    updateSession.run(1_021_000, 8_000, 5_000, 3, 30, 6, 5, now, s3);

    const animeDetail = getAnimeDetail(db, animeId);
    assert.ok(animeDetail);
    assert.equal(animeDetail?.totalLookupCount, 15);
    assert.equal(animeDetail?.totalLookupHits, 12);

    const episodes = getAnimeEpisodes(db, animeId);
    assert.deepEqual(
      episodes.map((row) => ({
        videoId: row.videoId,
        endedMediaMs: row.endedMediaMs,
        totalSessions: row.totalSessions,
        totalActiveMs: row.totalActiveMs,
        totalCards: row.totalCards,
        totalTokensSeen: row.totalTokensSeen,
      })),
      [
        {
          videoId: episodeOne,
          endedMediaMs: 6_000,
          totalSessions: 2,
          totalActiveMs: 7_000,
          totalCards: 3,
          totalTokensSeen: 30,
        },
        {
          videoId: episodeTwo,
          endedMediaMs: 8_000,
          totalSessions: 1,
          totalActiveMs: 5_000,
          totalCards: 3,
          totalTokensSeen: 30,
        },
      ],
    );

    const mediaDetail = getMediaDetail(db, episodeOne);
    assert.ok(mediaDetail);
    assert.equal(mediaDetail?.totalSessions, 2);
    assert.equal(mediaDetail?.totalActiveMs, 7_000);
    assert.equal(mediaDetail?.totalCards, 3);
    assert.equal(mediaDetail?.totalTokensSeen, 30);
    assert.equal(mediaDetail?.totalLookupCount, 9);
    assert.equal(mediaDetail?.totalLookupHits, 7);
    assert.equal(mediaDetail?.totalYomitanLookupCount, 0);
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
        sourcePath: '/tmp/Little Witch Academia S02E04.mkv',
        secondaryText: null,
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
        sourcePath: '/tmp/[SubsPlease] Frieren - 03 - Departure.mkv',
        secondaryText: null,
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

test('deleteSession removes the session and all associated session-scoped rows', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/delete-session.mkv', {
      canonicalTitle: 'Delete Session Test',
      sourcePath: '/tmp/delete-session.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const startedAtMs = 6_000_000;
    const { sessionId } = startSessionRecord(db, videoId, startedAtMs);

    stmts.telemetryInsertStmt.run(
      sessionId,
      startedAtMs + 1_000,
      5_000,
      4_000,
      3,
      9,
      9,
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
    const eventResult = stmts.eventInsertStmt.run(
      sessionId,
      startedAtMs + 1_500,
      EVENT_SUBTITLE_LINE,
      0,
      0,
      900,
      2,
      0,
      '{"line":"delete me"}',
      startedAtMs + 1_500,
      startedAtMs + 1_500,
    );
    const eventId = Number(eventResult.lastInsertRowid);
    const wordResult = db
      .prepare(
        `INSERT INTO imm_words (
          headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run('削除', '削除', 'さくじょ', 'noun', '名詞', '一般', '', startedAtMs, startedAtMs, 1);
    const kanjiResult = db
      .prepare(
        `INSERT INTO imm_kanji (
          kanji, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?)`,
      )
      .run('削', startedAtMs, startedAtMs, 1);
    const lineResult = stmts.subtitleLineInsertStmt.run(
      sessionId,
      eventId,
      videoId,
      null,
      0,
      0,
      900,
      'delete me',
      startedAtMs + 1_500,
      startedAtMs + 1_500,
    );
    const lineId = Number(lineResult.lastInsertRowid);
    db.prepare(
      `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(lineId, Number(wordResult.lastInsertRowid), 1);
    db.prepare(
      `INSERT INTO imm_kanji_line_occurrences (line_id, kanji_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(lineId, Number(kanjiResult.lastInsertRowid), 1);

    deleteSession(db, sessionId);

    const sessionCount = Number(
      (
        db
          .prepare('SELECT COUNT(*) AS total FROM imm_sessions WHERE session_id = ?')
          .get(sessionId) as {
          total: number;
        }
      ).total,
    );
    const telemetryCount = Number(
      (
        db
          .prepare('SELECT COUNT(*) AS total FROM imm_session_telemetry WHERE session_id = ?')
          .get(sessionId) as { total: number }
      ).total,
    );
    const eventCount = Number(
      (
        db
          .prepare('SELECT COUNT(*) AS total FROM imm_session_events WHERE session_id = ?')
          .get(sessionId) as {
          total: number;
        }
      ).total,
    );
    const subtitleLineCount = Number(
      (
        db
          .prepare('SELECT COUNT(*) AS total FROM imm_subtitle_lines WHERE session_id = ?')
          .get(sessionId) as { total: number }
      ).total,
    );
    const wordOccurrenceCount = Number(
      (
        db
          .prepare('SELECT COUNT(*) AS total FROM imm_word_line_occurrences WHERE line_id = ?')
          .get(lineId) as { total: number }
      ).total,
    );
    const kanjiOccurrenceCount = Number(
      (
        db
          .prepare('SELECT COUNT(*) AS total FROM imm_kanji_line_occurrences WHERE line_id = ?')
          .get(lineId) as { total: number }
      ).total,
    );

    assert.equal(sessionCount, 0);
    assert.equal(telemetryCount, 0);
    assert.equal(eventCount, 0);
    assert.equal(subtitleLineCount, 0);
    assert.equal(wordOccurrenceCount, 0);
    assert.equal(kanjiOccurrenceCount, 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('deleteSession rebuilds word and kanji aggregates from retained subtitle lines', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/delete-session-aggregates.mkv', {
      canonicalTitle: 'Delete Session Aggregates Test',
      sourcePath: '/tmp/delete-session-aggregates.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const deletedSession = startSessionRecord(db, videoId, 7_000_000);
    const keptSession = startSessionRecord(db, videoId, 8_000_000);
    const deletedTs = 7_000_500;
    const keptTs = 8_000_500;

    const sharedWordResult = db
      .prepare(
        `INSERT INTO imm_words (
          headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run('共有', '共有', 'きょうゆう', 'noun', '名詞', '一般', '', deletedTs, keptTs, 3);
    const deletedOnlyWordResult = db
      .prepare(
        `INSERT INTO imm_words (
          headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run(
        '削除専用',
        '削除専用',
        'さくじょせんよう',
        'noun',
        '名詞',
        '一般',
        '',
        deletedTs,
        deletedTs,
        1,
      );
    const sharedKanjiResult = db
      .prepare(
        `INSERT INTO imm_kanji (
          kanji, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?)`,
      )
      .run('共', deletedTs, keptTs, 3);
    const deletedOnlyKanjiResult = db
      .prepare(
        `INSERT INTO imm_kanji (
          kanji, first_seen, last_seen, frequency
        ) VALUES (?, ?, ?, ?)`,
      )
      .run('削', deletedTs, deletedTs, 1);

    const deletedLineResult = stmts.subtitleLineInsertStmt.run(
      deletedSession.sessionId,
      null,
      videoId,
      null,
      0,
      0,
      800,
      'delete me',
      deletedTs,
      deletedTs,
    );
    const keptLineResult = stmts.subtitleLineInsertStmt.run(
      keptSession.sessionId,
      null,
      videoId,
      null,
      0,
      1_000,
      1_800,
      'keep me',
      keptTs,
      keptTs,
    );

    const deletedLineId = Number(deletedLineResult.lastInsertRowid);
    const keptLineId = Number(keptLineResult.lastInsertRowid);
    const sharedWordId = Number(sharedWordResult.lastInsertRowid);
    const deletedOnlyWordId = Number(deletedOnlyWordResult.lastInsertRowid);
    const sharedKanjiId = Number(sharedKanjiResult.lastInsertRowid);
    const deletedOnlyKanjiId = Number(deletedOnlyKanjiResult.lastInsertRowid);

    db.prepare(
      `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(deletedLineId, sharedWordId, 2);
    db.prepare(
      `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(deletedLineId, deletedOnlyWordId, 1);
    db.prepare(
      `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(keptLineId, sharedWordId, 1);
    db.prepare(
      `INSERT INTO imm_kanji_line_occurrences (line_id, kanji_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(deletedLineId, sharedKanjiId, 2);
    db.prepare(
      `INSERT INTO imm_kanji_line_occurrences (line_id, kanji_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(deletedLineId, deletedOnlyKanjiId, 1);
    db.prepare(
      `INSERT INTO imm_kanji_line_occurrences (line_id, kanji_id, occurrence_count)
       VALUES (?, ?, ?)`,
    ).run(keptLineId, sharedKanjiId, 1);

    deleteSession(db, deletedSession.sessionId);

    const sharedWordRow = db
      .prepare('SELECT frequency, first_seen, last_seen FROM imm_words WHERE id = ?')
      .get(sharedWordId) as {
      frequency: number;
      first_seen: number;
      last_seen: number;
    } | null;
    const deletedOnlyWordRow = db
      .prepare('SELECT id FROM imm_words WHERE id = ?')
      .get(deletedOnlyWordId) as { id: number } | null;
    const sharedKanjiRow = db
      .prepare('SELECT frequency, first_seen, last_seen FROM imm_kanji WHERE id = ?')
      .get(sharedKanjiId) as {
      frequency: number;
      first_seen: number;
      last_seen: number;
    } | null;
    const deletedOnlyKanjiRow = db
      .prepare('SELECT id FROM imm_kanji WHERE id = ?')
      .get(deletedOnlyKanjiId) as { id: number } | null;

    assert.ok(sharedWordRow);
    assert.equal(sharedWordRow.frequency, 1);
    assert.equal(sharedWordRow.first_seen, keptTs);
    assert.equal(sharedWordRow.last_seen, keptTs);
    assert.equal(deletedOnlyWordRow ?? null, null);
    assert.ok(sharedKanjiRow);
    assert.equal(sharedKanjiRow.frequency, 1);
    assert.equal(sharedKanjiRow.first_seen, keptTs);
    assert.equal(sharedKanjiRow.last_seen, keptTs);
    assert.equal(deletedOnlyKanjiRow ?? null, null);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});
