import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Database } from '../sqlite.js';
import type { DatabaseSync } from '../sqlite.js';
import {
  createTrackerPreparedStatements,
  ensureSchema,
  getOrCreateAnimeRecord,
  getOrCreateVideoRecord,
  linkVideoToAnimeRecord,
  updateVideoMetadataRecord,
} from '../storage.js';
import { startSessionRecord } from '../session.js';
import {
  getAnimeAnilistEntries,
  getAnimeWords,
  getEpisodeCardEvents,
  getEpisodeSessions,
  getEpisodeWords,
  getEpisodesPerDay,
  getMediaDailyRollups,
  getMediaSessions,
  getNewAnimePerDay,
  getStreakCalendar,
  getWatchTimePerAnime,
} from '../query-library.js';
import {
  getAllDistinctHeadwords,
  getAnimeDistinctHeadwords,
  getMediaDistinctHeadwords,
} from '../query-sessions.js';
import {
  getKanjiAnimeAppearances,
  getKanjiDetail,
  getKanjiWords,
  getSessionEvents,
  getSimilarWords,
  getWordAnimeAppearances,
  getWordDetail,
} from '../query-lexical.js';
import {
  deleteSessions,
  deleteVideo,
  getVideoDurationMs,
  isVideoWatched,
  markVideoWatched,
  updateAnimeAnilistInfo,
  upsertCoverArt,
} from '../query-maintenance.js';
import { getLocalEpochDay } from '../query-shared.js';
import { EVENT_CARD_MINED, EVENT_SUBTITLE_LINE, SOURCE_TYPE_LOCAL } from '../types.js';

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-imm-query-split-test-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  const dir = path.dirname(dbPath);
  if (!fs.existsSync(dir)) return;
  fs.rmSync(dir, { recursive: true, force: true });
}

function createDb() {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  ensureSchema(db);
  const stmts = createTrackerPreparedStatements(db);
  return { db, dbPath, stmts };
}

function finalizeSessionMetrics(
  db: DatabaseSync,
  sessionId: number,
  startedAtMs: number,
  options: {
    endedAtMs?: number;
    totalWatchedMs?: number;
    activeWatchedMs?: number;
    linesSeen?: number;
    tokensSeen?: number;
    cardsMined?: number;
    lookupCount?: number;
    lookupHits?: number;
    yomitanLookupCount?: number;
  } = {},
): void {
  const endedAtMs = options.endedAtMs ?? startedAtMs + 60_000;
  db.prepare(
    `
      UPDATE imm_sessions
      SET
        ended_at_ms = ?,
        status = 2,
        ended_media_ms = ?,
        total_watched_ms = ?,
        active_watched_ms = ?,
        lines_seen = ?,
        tokens_seen = ?,
        cards_mined = ?,
        lookup_count = ?,
        lookup_hits = ?,
        yomitan_lookup_count = ?,
        LAST_UPDATE_DATE = ?
      WHERE session_id = ?
    `,
  ).run(
    endedAtMs,
    options.totalWatchedMs ?? 50_000,
    options.totalWatchedMs ?? 50_000,
    options.activeWatchedMs ?? 45_000,
    options.linesSeen ?? 3,
    options.tokensSeen ?? 6,
    options.cardsMined ?? 1,
    options.lookupCount ?? 2,
    options.lookupHits ?? 1,
    options.yomitanLookupCount ?? 1,
    endedAtMs,
    sessionId,
  );
}

function insertWordOccurrence(
  db: DatabaseSync,
  stmts: ReturnType<typeof createTrackerPreparedStatements>,
  options: {
    sessionId: number;
    videoId: number;
    animeId: number | null;
    lineIndex: number;
    text: string;
    word: { headword: string; word: string; reading: string; pos?: string };
    occurrenceCount?: number;
  },
): number {
  const nowMs = 1_000_000 + options.lineIndex;
  stmts.wordUpsertStmt.run(
    options.word.headword,
    options.word.word,
    options.word.reading,
    options.word.pos ?? 'noun',
    '名詞',
    '一般',
    '',
    nowMs,
    nowMs,
  );
  const wordRow = db
    .prepare('SELECT id FROM imm_words WHERE headword = ? AND word = ? AND reading = ?')
    .get(options.word.headword, options.word.word, options.word.reading) as { id: number };
  const lineResult = stmts.subtitleLineInsertStmt.run(
    options.sessionId,
    null,
    options.videoId,
    options.animeId,
    options.lineIndex,
    options.lineIndex * 1000,
    options.lineIndex * 1000 + 900,
    options.text,
    '',
    nowMs,
    nowMs,
  );
  const lineId = Number(lineResult.lastInsertRowid);
  stmts.wordLineOccurrenceUpsertStmt.run(lineId, wordRow.id, options.occurrenceCount ?? 1);
  return wordRow.id;
}

function insertKanjiOccurrence(
  db: DatabaseSync,
  stmts: ReturnType<typeof createTrackerPreparedStatements>,
  options: {
    sessionId: number;
    videoId: number;
    animeId: number | null;
    lineIndex: number;
    text: string;
    kanji: string;
    occurrenceCount?: number;
  },
): number {
  const nowMs = 2_000_000 + options.lineIndex;
  stmts.kanjiUpsertStmt.run(options.kanji, nowMs, nowMs);
  const kanjiRow = db.prepare('SELECT id FROM imm_kanji WHERE kanji = ?').get(options.kanji) as {
    id: number;
  };
  const lineResult = stmts.subtitleLineInsertStmt.run(
    options.sessionId,
    null,
    options.videoId,
    options.animeId,
    options.lineIndex,
    options.lineIndex * 1000,
    options.lineIndex * 1000 + 900,
    options.text,
    '',
    nowMs,
    nowMs,
  );
  const lineId = Number(lineResult.lastInsertRowid);
  stmts.kanjiLineOccurrenceUpsertStmt.run(lineId, kanjiRow.id, options.occurrenceCount ?? 1);
  return kanjiRow.id;
}

test('split session and lexical helpers return distinct-headword, detail, appearance, and filter results', () => {
  const { db, dbPath, stmts } = createDb();

  try {
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Lexical Anime',
      canonicalTitle: 'Lexical Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/lexical-episode-1.mkv', {
      canonicalTitle: 'Lexical Episode 1',
      sourcePath: '/tmp/lexical-episode-1.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: 'lexical-episode-1.mkv',
      parsedTitle: 'Lexical Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'test',
      parserConfidence: 1,
      parseMetadataJson: null,
    });
    const sessionId = startSessionRecord(db, videoId, 1_000_000).sessionId;

    const nekoId = insertWordOccurrence(db, stmts, {
      sessionId,
      videoId,
      animeId,
      lineIndex: 1,
      text: '猫がいる',
      word: { headword: '猫', word: '猫', reading: 'ねこ' },
      occurrenceCount: 2,
    });
    insertWordOccurrence(db, stmts, {
      sessionId,
      videoId,
      animeId,
      lineIndex: 2,
      text: '犬もいる',
      word: { headword: '犬', word: '犬', reading: 'いぬ' },
    });
    insertWordOccurrence(db, stmts, {
      sessionId,
      videoId,
      animeId,
      lineIndex: 3,
      text: '子猫だ',
      word: { headword: '子猫', word: '子猫', reading: 'こねこ' },
    });
    insertWordOccurrence(db, stmts, {
      sessionId,
      videoId,
      animeId,
      lineIndex: 5,
      text: '日本だ',
      word: { headword: '日本', word: '日本', reading: 'にほん' },
    });
    const hiId = insertKanjiOccurrence(db, stmts, {
      sessionId,
      videoId,
      animeId,
      lineIndex: 4,
      text: '日本',
      kanji: '日',
      occurrenceCount: 3,
    });

    stmts.eventInsertStmt.run(
      sessionId,
      1_000_100,
      EVENT_SUBTITLE_LINE,
      1,
      0,
      900,
      0,
      0,
      JSON.stringify({ kind: 'subtitle' }),
      1_000_100,
      1_000_100,
    );
    stmts.eventInsertStmt.run(
      sessionId,
      1_000_200,
      EVENT_CARD_MINED,
      2,
      1000,
      1900,
      0,
      1,
      JSON.stringify({ noteIds: [41] }),
      1_000_200,
      1_000_200,
    );

    assert.deepEqual(getAllDistinctHeadwords(db).sort(), ['子猫', '日本', '犬', '猫']);
    assert.deepEqual(getAnimeDistinctHeadwords(db, animeId).sort(), ['子猫', '日本', '犬', '猫']);
    assert.deepEqual(getMediaDistinctHeadwords(db, videoId).sort(), ['子猫', '日本', '犬', '猫']);

    const wordDetail = getWordDetail(db, nekoId);
    assert.ok(wordDetail);
    assert.equal(wordDetail.wordId, nekoId);
    assert.equal(wordDetail.headword, '猫');
    assert.equal(wordDetail.word, '猫');
    assert.equal(wordDetail.reading, 'ねこ');
    assert.equal(wordDetail.partOfSpeech, 'noun');
    assert.equal(wordDetail.pos1, '名詞');
    assert.equal(wordDetail.pos2, '一般');
    assert.equal(wordDetail.pos3, '');
    assert.equal(wordDetail.frequency, 1);
    assert.equal(wordDetail.firstSeen, 1_000_001);
    assert.equal(wordDetail.lastSeen, 1_000_001);
    assert.deepEqual(getWordAnimeAppearances(db, nekoId), [
      { animeId, animeTitle: 'Lexical Anime', occurrenceCount: 2 },
    ]);
    assert.deepEqual(
      getSimilarWords(db, nekoId, 5).map((row) => row.headword),
      ['子猫'],
    );

    const kanjiDetail = getKanjiDetail(db, hiId);
    assert.ok(kanjiDetail);
    assert.equal(kanjiDetail.kanjiId, hiId);
    assert.equal(kanjiDetail.kanji, '日');
    assert.equal(kanjiDetail.frequency, 1);
    assert.equal(kanjiDetail.firstSeen, 2_000_004);
    assert.equal(kanjiDetail.lastSeen, 2_000_004);
    assert.deepEqual(getKanjiAnimeAppearances(db, hiId), [
      { animeId, animeTitle: 'Lexical Anime', occurrenceCount: 3 },
    ]);
    assert.deepEqual(
      getKanjiWords(db, hiId, 5).map((row) => row.headword),
      ['日本'],
    );

    assert.deepEqual(
      getSessionEvents(db, sessionId, 10, [EVENT_CARD_MINED]).map((row) => row.eventType),
      [EVENT_CARD_MINED],
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('split library helpers return anime/media session and analytics rows', () => {
  const { db, dbPath, stmts } = createDb();

  try {
    const now = new Date();
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Library Anime',
      canonicalTitle: 'Library Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/library-episode-1.mkv', {
      canonicalTitle: 'Library Episode 1',
      sourcePath: '/tmp/library-episode-1.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: 'library-episode-1.mkv',
      parsedTitle: 'Library Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'test',
      parserConfidence: 1,
      parseMetadataJson: null,
    });

    const startedAtMs = new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
      9,
      0,
      0,
    ).getTime();
    const sessionId = startSessionRecord(db, videoId, startedAtMs).sessionId;
    const todayLocalDay = getLocalEpochDay(db, startedAtMs);
    finalizeSessionMetrics(db, sessionId, startedAtMs, {
      endedAtMs: startedAtMs + 55_000,
      totalWatchedMs: 55_000,
      activeWatchedMs: 45_000,
      linesSeen: 4,
      tokensSeen: 8,
      cardsMined: 2,
    });
    db.prepare(
      `
        INSERT INTO imm_daily_rollups (
          rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
          total_tokens_seen, total_cards, cards_per_hour, tokens_per_min, lookup_hit_rate,
          CREATED_DATE, LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(todayLocalDay, videoId, 1, 45, 4, 8, 2, 2.66, 0.17, 0.5, startedAtMs, startedAtMs);

    db.prepare(
      `
        INSERT INTO imm_media_art (
          video_id, anilist_id, cover_url, cover_blob, cover_blob_hash, title_romaji,
          title_english, episodes_total, fetched_at_ms, CREATED_DATE, LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(
      videoId,
      77,
      'https://images.test/library.jpg',
      new Uint8Array([1, 2, 3]),
      null,
      'Library Anime',
      'Library Anime',
      12,
      startedAtMs,
      startedAtMs,
      startedAtMs,
    );

    db.prepare(
      `
        INSERT INTO imm_session_events (
          session_id, ts_ms, event_type, line_index, segment_start_ms, segment_end_ms,
          tokens_delta, cards_delta, payload_json, CREATED_DATE, LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(
      sessionId,
      startedAtMs + 40_000,
      EVENT_CARD_MINED,
      4,
      4000,
      4900,
      0,
      2,
      JSON.stringify({ noteIds: [101, 102] }),
      startedAtMs + 40_000,
      startedAtMs + 40_000,
    );

    insertWordOccurrence(db, stmts, {
      sessionId,
      videoId,
      animeId,
      lineIndex: 1,
      text: '猫がいる',
      word: { headword: '猫', word: '猫', reading: 'ねこ' },
      occurrenceCount: 3,
    });
    insertWordOccurrence(db, stmts, {
      sessionId,
      videoId,
      animeId,
      lineIndex: 2,
      text: '犬もいる',
      word: { headword: '犬', word: '犬', reading: 'いぬ' },
      occurrenceCount: 1,
    });

    assert.deepEqual(getAnimeAnilistEntries(db, animeId), [
      {
        anilistId: 77,
        titleRomaji: 'Library Anime',
        titleEnglish: 'Library Anime',
        season: 1,
      },
    ]);
    assert.equal(getMediaSessions(db, videoId, 10)[0]?.sessionId, sessionId);
    assert.equal(getEpisodeSessions(db, videoId)[0]?.sessionId, sessionId);
    assert.equal(getMediaDailyRollups(db, videoId, 10)[0]?.totalActiveMin, 45);
    assert.deepEqual(getStreakCalendar(db, 30), [{ epochDay: todayLocalDay, totalActiveMin: 45 }]);
    assert.deepEqual(
      getAnimeWords(db, animeId, 10).map((row) => row.headword),
      ['猫', '犬'],
    );
    assert.deepEqual(
      getEpisodeWords(db, videoId, 10).map((row) => row.headword),
      ['猫', '犬'],
    );
    assert.deepEqual(getEpisodesPerDay(db, 10), [{ epochDay: todayLocalDay, episodeCount: 1 }]);
    assert.deepEqual(getNewAnimePerDay(db, 10), [{ epochDay: todayLocalDay, newAnimeCount: 1 }]);
    assert.deepEqual(getWatchTimePerAnime(db, 3650), [
      {
        epochDay: todayLocalDay,
        animeId,
        animeTitle: 'Library Anime',
        totalActiveMin: 45,
      },
    ]);
    assert.deepEqual(getEpisodeCardEvents(db, videoId), [
      {
        eventId: 1,
        sessionId,
        tsMs: startedAtMs + 40_000,
        cardsDelta: 2,
        noteIds: [101, 102],
      },
    ]);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('split maintenance helpers update anime metadata and watched state', () => {
  const { db, dbPath } = createDb();

  try {
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Metadata Anime',
      canonicalTitle: 'Metadata Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/metadata-episode-1.mkv', {
      canonicalTitle: 'Metadata Episode 1',
      sourcePath: '/tmp/metadata-episode-1.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: 'metadata-episode-1.mkv',
      parsedTitle: 'Metadata Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'test',
      parserConfidence: 1,
      parseMetadataJson: null,
    });
    updateVideoMetadataRecord(db, videoId, {
      sourceType: SOURCE_TYPE_LOCAL,
      canonicalTitle: 'Metadata Episode 1',
      durationMs: 222_000,
      fileSizeBytes: null,
      codecId: null,
      containerId: null,
      widthPx: null,
      heightPx: null,
      fpsX100: null,
      bitrateKbps: null,
      audioCodecId: null,
      hashSha256: null,
      screenshotPath: null,
      metadataJson: null,
    });

    updateAnimeAnilistInfo(db, videoId, {
      anilistId: 99,
      titleRomaji: 'Metadata Romaji',
      titleEnglish: 'Metadata English',
      titleNative: 'メタデータ',
      episodesTotal: 24,
    });
    markVideoWatched(db, videoId, true);

    const animeRow = db
      .prepare(
        `
          SELECT anilist_id, title_romaji, title_english, title_native, episodes_total
          FROM imm_anime
          WHERE anime_id = ?
        `,
      )
      .get(animeId) as {
      anilist_id: number;
      title_romaji: string;
      title_english: string;
      title_native: string;
      episodes_total: number;
    };

    assert.equal(animeRow.anilist_id, 99);
    assert.equal(animeRow.title_romaji, 'Metadata Romaji');
    assert.equal(animeRow.title_english, 'Metadata English');
    assert.equal(animeRow.title_native, 'メタデータ');
    assert.equal(animeRow.episodes_total, 24);
    assert.equal(getVideoDurationMs(db, videoId), 222_000);
    assert.equal(isVideoWatched(db, videoId), true);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('split maintenance helpers delete multiple sessions and whole videos with dependent rows', () => {
  const { db, dbPath, stmts } = createDb();

  try {
    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Delete Anime',
      canonicalTitle: 'Delete Anime',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    const keepVideoId = getOrCreateVideoRecord(db, 'local:/tmp/delete-keep.mkv', {
      canonicalTitle: 'Delete Keep',
      sourcePath: '/tmp/delete-keep.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const dropVideoId = getOrCreateVideoRecord(db, 'local:/tmp/delete-drop.mkv', {
      canonicalTitle: 'Delete Drop',
      sourcePath: '/tmp/delete-drop.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    linkVideoToAnimeRecord(db, keepVideoId, {
      animeId,
      parsedBasename: 'delete-keep.mkv',
      parsedTitle: 'Delete Anime',
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'test',
      parserConfidence: 1,
      parseMetadataJson: null,
    });
    linkVideoToAnimeRecord(db, dropVideoId, {
      animeId,
      parsedBasename: 'delete-drop.mkv',
      parsedTitle: 'Delete Anime',
      parsedSeason: 1,
      parsedEpisode: 2,
      parserSource: 'test',
      parserConfidence: 1,
      parseMetadataJson: null,
    });

    const keepSessionId = startSessionRecord(db, keepVideoId, 1_000_000).sessionId;
    const dropSessionOne = startSessionRecord(db, dropVideoId, 2_000_000).sessionId;
    const dropSessionTwo = startSessionRecord(db, dropVideoId, 3_000_000).sessionId;
    finalizeSessionMetrics(db, keepSessionId, 1_000_000);
    finalizeSessionMetrics(db, dropSessionOne, 2_000_000);
    finalizeSessionMetrics(db, dropSessionTwo, 3_000_000);

    insertWordOccurrence(db, stmts, {
      sessionId: dropSessionOne,
      videoId: dropVideoId,
      animeId,
      lineIndex: 1,
      text: '削除する猫',
      word: { headword: '猫', word: '猫', reading: 'ねこ' },
    });
    insertKanjiOccurrence(db, stmts, {
      sessionId: dropSessionOne,
      videoId: dropVideoId,
      animeId,
      lineIndex: 2,
      text: '日本',
      kanji: '日',
    });
    upsertCoverArt(db, dropVideoId, {
      anilistId: 12,
      coverUrl: 'https://images.test/delete.jpg',
      coverBlob: new Uint8Array([7, 8, 9]),
      titleRomaji: 'Delete Anime',
      titleEnglish: 'Delete Anime',
      episodesTotal: 2,
    });

    deleteSessions(db, [dropSessionOne, dropSessionTwo]);

    const deletedSessionCount = db
      .prepare('SELECT COUNT(*) AS total FROM imm_sessions WHERE video_id = ?')
      .get(dropVideoId) as { total: number };
    assert.equal(deletedSessionCount.total, 0);

    const keepReplacementSession = startSessionRecord(db, keepVideoId, 4_000_000).sessionId;
    finalizeSessionMetrics(db, keepReplacementSession, 4_000_000);

    deleteVideo(db, dropVideoId);

    const remainingVideos = db
      .prepare('SELECT video_id FROM imm_videos ORDER BY video_id')
      .all() as Array<{
      video_id: number;
    }>;
    const coverRows = db.prepare('SELECT COUNT(*) AS total FROM imm_media_art').get() as {
      total: number;
    };

    assert.deepEqual(remainingVideos, [{ video_id: keepVideoId }]);
    assert.equal(coverRows.total, 0);
    assert.equal(
      (
        db.prepare('SELECT COUNT(*) AS total FROM imm_words').get() as {
          total: number;
        }
      ).total,
      0,
    );
    assert.equal(
      (
        db.prepare('SELECT COUNT(*) AS total FROM imm_kanji').get() as {
          total: number;
        }
      ).total,
      0,
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});
