import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Database } from './sqlite';
import { getStatsExcludedWords, replaceStatsExcludedWords } from './query-lexical';
import { finalizeSessionRecord, startSessionRecord } from './session';
import {
  applyPragmas,
  createTrackerPreparedStatements,
  ensureSchema,
  executeQueuedWrite,
  normalizeCoverBlobBytes,
  parseCoverBlobReference,
  getOrCreateAnimeRecord,
  getOrCreateVideoRecord,
  linkVideoToAnimeRecord,
  linkYoutubeVideoToAnimeRecord,
} from './storage';
import {
  EVENT_SUBTITLE_LINE,
  SESSION_STATUS_ENDED,
  SOURCE_TYPE_LOCAL,
  SOURCE_TYPE_REMOTE,
} from './types';

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-imm-storage-session-'));
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

test('applyPragmas sets the SQLite tuning defaults used by immersion tracking', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    applyPragmas(db);

    const journalModeRow = db.prepare('PRAGMA journal_mode').get() as {
      journal_mode: string;
    };
    const synchronousRow = db.prepare('PRAGMA synchronous').get() as { synchronous: number };
    const foreignKeysRow = db.prepare('PRAGMA foreign_keys').get() as { foreign_keys: number };
    const busyTimeoutRow = db.prepare('PRAGMA busy_timeout').get() as { timeout: number };
    const journalSizeLimitRow = db.prepare('PRAGMA journal_size_limit').get() as {
      journal_size_limit: number;
    };

    assert.equal(journalModeRow.journal_mode, 'wal');
    assert.equal(synchronousRow.synchronous, 1);
    assert.equal(foreignKeysRow.foreign_keys, 1);
    assert.equal(busyTimeoutRow.timeout, 2500);
    assert.equal(journalSizeLimitRow.journal_size_limit, 67_108_864);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('ensureSchema creates immersion core tables', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const rows = db
      .prepare(
        `SELECT name FROM sqlite_master WHERE type = 'table' AND name LIKE 'imm_%' ORDER BY name`,
      )
      .all() as Array<{ name: string }>;
    const tableNames = new Set(rows.map((row) => row.name));

    assert.ok(tableNames.has('imm_videos'));
    assert.ok(tableNames.has('imm_anime'));
    assert.ok(tableNames.has('imm_sessions'));
    assert.ok(tableNames.has('imm_session_telemetry'));
    assert.ok(tableNames.has('imm_session_events'));
    assert.ok(tableNames.has('imm_daily_rollups'));
    assert.ok(tableNames.has('imm_monthly_rollups'));
    assert.ok(tableNames.has('imm_words'));
    assert.ok(tableNames.has('imm_kanji'));
    assert.ok(tableNames.has('imm_subtitle_lines'));
    assert.ok(tableNames.has('imm_word_line_occurrences'));
    assert.ok(tableNames.has('imm_kanji_line_occurrences'));
    assert.ok(tableNames.has('imm_rollup_state'));
    assert.ok(tableNames.has('imm_cover_art_blobs'));
    assert.ok(tableNames.has('imm_youtube_videos'));
    assert.ok(tableNames.has('imm_stats_excluded_words'));

    const videoColumns = new Set(
      (
        db.prepare('PRAGMA table_info(imm_videos)').all() as Array<{
          name: string;
        }>
      ).map((row) => row.name),
    );

    assert.ok(videoColumns.has('anime_id'));
    assert.ok(videoColumns.has('parsed_basename'));
    assert.ok(videoColumns.has('parsed_title'));
    assert.ok(videoColumns.has('parsed_season'));
    assert.ok(videoColumns.has('parsed_episode'));
    assert.ok(videoColumns.has('parser_source'));
    assert.ok(videoColumns.has('parser_confidence'));
    assert.ok(videoColumns.has('parse_metadata_json'));

    const mediaArtColumns = new Set(
      (
        db.prepare('PRAGMA table_info(imm_media_art)').all() as Array<{
          name: string;
        }>
      ).map((row) => row.name),
    );
    assert.ok(mediaArtColumns.has('cover_blob_hash'));

    const rollupStateRow = db
      .prepare('SELECT state_value FROM imm_rollup_state WHERE state_key = ?')
      .get('last_rollup_sample_ms') as {
      state_value: string;
    } | null;
    assert.ok(rollupStateRow);
    assert.equal(Number(rollupStateRow?.state_value ?? 0), 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('stats excluded words are replaced and read from sqlite storage', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    replaceStatsExcludedWords(db, [
      { headword: '猫', word: '猫', reading: 'ねこ' },
      { headword: 'する', word: 'する', reading: 'する' },
    ]);
    assert.deepEqual(getStatsExcludedWords(db), [
      { headword: 'する', word: 'する', reading: 'する' },
      { headword: '猫', word: '猫', reading: 'ねこ' },
    ]);

    replaceStatsExcludedWords(db, [{ headword: '犬', word: '犬', reading: 'いぬ' }]);
    assert.deepEqual(getStatsExcludedWords(db), [{ headword: '犬', word: '犬', reading: 'いぬ' }]);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('ensureSchema adds youtube metadata table to existing schema version 15 databases', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    db.exec(`
      CREATE TABLE imm_schema_version (
        schema_version INTEGER PRIMARY KEY,
        applied_at_ms INTEGER NOT NULL
      );
      INSERT INTO imm_schema_version(schema_version, applied_at_ms) VALUES (15, 1000);

      CREATE TABLE imm_rollup_state(
        state_key TEXT PRIMARY KEY,
        state_value INTEGER NOT NULL
      );
      INSERT INTO imm_rollup_state(state_key, state_value) VALUES ('last_rollup_sample_ms', 123);

      CREATE TABLE imm_anime(
        anime_id INTEGER PRIMARY KEY AUTOINCREMENT,
        normalized_title_key TEXT NOT NULL UNIQUE,
        canonical_title TEXT NOT NULL,
        anilist_id INTEGER UNIQUE,
        title_romaji TEXT,
        title_english TEXT,
        title_native TEXT,
        episodes_total INTEGER,
        description TEXT,
        metadata_json TEXT,
        CREATED_DATE INTEGER,
        LAST_UPDATE_DATE INTEGER
      );

      CREATE TABLE imm_videos(
        video_id INTEGER PRIMARY KEY AUTOINCREMENT,
        video_key TEXT NOT NULL UNIQUE,
        anime_id INTEGER,
        canonical_title TEXT NOT NULL,
        source_type INTEGER NOT NULL,
        source_path TEXT,
        source_url TEXT,
        parsed_basename TEXT,
        parsed_title TEXT,
        parsed_season INTEGER,
        parsed_episode INTEGER,
        parser_source TEXT,
        parser_confidence REAL,
        parse_metadata_json TEXT,
        watched INTEGER NOT NULL DEFAULT 0,
        duration_ms INTEGER NOT NULL CHECK(duration_ms>=0),
        file_size_bytes INTEGER CHECK(file_size_bytes>=0),
        codec_id INTEGER, container_id INTEGER,
        width_px INTEGER, height_px INTEGER, fps_x100 INTEGER,
        bitrate_kbps INTEGER, audio_codec_id INTEGER,
        hash_sha256 TEXT, screenshot_path TEXT,
        metadata_json TEXT,
        CREATED_DATE INTEGER,
        LAST_UPDATE_DATE INTEGER,
        FOREIGN KEY(anime_id) REFERENCES imm_anime(anime_id) ON DELETE SET NULL
      );
    `);

    ensureSchema(db);

    const tables = new Set(
      (
        db
          .prepare(`SELECT name FROM sqlite_master WHERE type = 'table' AND name LIKE 'imm_%'`)
          .all() as Array<{
          name: string;
        }>
      ).map((row) => row.name),
    );
    assert.ok(tables.has('imm_youtube_videos'));

    const columns = new Set(
      (
        db.prepare('PRAGMA table_info(imm_youtube_videos)').all() as Array<{
          name: string;
        }>
      ).map((row) => row.name),
    );

    assert.deepEqual(
      columns,
      new Set([
        'video_id',
        'youtube_video_id',
        'video_url',
        'video_title',
        'video_thumbnail_url',
        'channel_id',
        'channel_name',
        'channel_url',
        'channel_thumbnail_url',
        'uploader_id',
        'uploader_url',
        'description',
        'metadata_json',
        'fetched_at_ms',
        'CREATED_DATE',
        'LAST_UPDATE_DATE',
      ]),
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('ensureSchema migrates session event timestamps to text and repairs libsql-truncated wall-clock values', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    db.exec(`
      CREATE TABLE imm_schema_version (
        schema_version INTEGER PRIMARY KEY,
        applied_at_ms INTEGER NOT NULL
      );
      INSERT INTO imm_schema_version(schema_version, applied_at_ms) VALUES (16, 1000);

      CREATE TABLE imm_rollup_state(
        state_key TEXT PRIMARY KEY,
        state_value INTEGER NOT NULL
      );
      INSERT INTO imm_rollup_state(state_key, state_value) VALUES ('last_rollup_sample_ms', 0);

      CREATE TABLE imm_anime(
        anime_id INTEGER PRIMARY KEY AUTOINCREMENT,
        normalized_title_key TEXT NOT NULL UNIQUE,
        canonical_title TEXT NOT NULL,
        anilist_id INTEGER UNIQUE,
        title_romaji TEXT,
        title_english TEXT,
        title_native TEXT,
        episodes_total INTEGER,
        description TEXT,
        metadata_json TEXT,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT
      );

      CREATE TABLE imm_videos(
        video_id INTEGER PRIMARY KEY AUTOINCREMENT,
        video_key TEXT NOT NULL UNIQUE,
        anime_id INTEGER,
        canonical_title TEXT NOT NULL,
        source_type INTEGER NOT NULL,
        source_path TEXT,
        source_url TEXT,
        parsed_basename TEXT,
        parsed_title TEXT,
        parsed_season INTEGER,
        parsed_episode INTEGER,
        parser_source TEXT,
        parser_confidence REAL,
        parse_metadata_json TEXT,
        watched INTEGER NOT NULL DEFAULT 0,
        duration_ms INTEGER NOT NULL CHECK(duration_ms>=0),
        file_size_bytes INTEGER CHECK(file_size_bytes>=0),
        codec_id INTEGER, container_id INTEGER,
        width_px INTEGER, height_px INTEGER, fps_x100 INTEGER,
        bitrate_kbps INTEGER, audio_codec_id INTEGER,
        hash_sha256 TEXT, screenshot_path TEXT,
        metadata_json TEXT,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        FOREIGN KEY(anime_id) REFERENCES imm_anime(anime_id) ON DELETE SET NULL
      );

      CREATE TABLE imm_sessions(
        session_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_uuid TEXT NOT NULL UNIQUE,
        video_id INTEGER NOT NULL,
        started_at_ms TEXT NOT NULL,
        ended_at_ms TEXT,
        status INTEGER NOT NULL,
        locale_id INTEGER,
        target_lang_id INTEGER,
        difficulty_tier INTEGER,
        subtitle_mode INTEGER,
        ended_media_ms INTEGER,
        total_watched_ms INTEGER NOT NULL DEFAULT 0,
        active_watched_ms INTEGER NOT NULL DEFAULT 0,
        lines_seen INTEGER NOT NULL DEFAULT 0,
        tokens_seen INTEGER NOT NULL DEFAULT 0,
        cards_mined INTEGER NOT NULL DEFAULT 0,
        lookup_count INTEGER NOT NULL DEFAULT 0,
        lookup_hits INTEGER NOT NULL DEFAULT 0,
        yomitan_lookup_count INTEGER NOT NULL DEFAULT 0,
        pause_count INTEGER NOT NULL DEFAULT 0,
        pause_ms INTEGER NOT NULL DEFAULT 0,
        seek_forward_count INTEGER NOT NULL DEFAULT 0,
        seek_backward_count INTEGER NOT NULL DEFAULT 0,
        media_buffer_events INTEGER NOT NULL DEFAULT 0,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        FOREIGN KEY(video_id) REFERENCES imm_videos(video_id)
      );

      CREATE TABLE imm_session_telemetry(
        telemetry_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        sample_ms TEXT NOT NULL,
        total_watched_ms INTEGER NOT NULL DEFAULT 0,
        active_watched_ms INTEGER NOT NULL DEFAULT 0,
        lines_seen INTEGER NOT NULL DEFAULT 0,
        tokens_seen INTEGER NOT NULL DEFAULT 0,
        cards_mined INTEGER NOT NULL DEFAULT 0,
        lookup_count INTEGER NOT NULL DEFAULT 0,
        lookup_hits INTEGER NOT NULL DEFAULT 0,
        yomitan_lookup_count INTEGER NOT NULL DEFAULT 0,
        pause_count INTEGER NOT NULL DEFAULT 0,
        pause_ms INTEGER NOT NULL DEFAULT 0,
        seek_forward_count INTEGER NOT NULL DEFAULT 0,
        seek_backward_count INTEGER NOT NULL DEFAULT 0,
        media_buffer_events INTEGER NOT NULL DEFAULT 0,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
      );

      CREATE TABLE imm_session_events(
        event_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        ts_ms INTEGER NOT NULL,
        event_type INTEGER NOT NULL,
        line_index INTEGER,
        segment_start_ms INTEGER,
        segment_end_ms INTEGER,
        tokens_delta INTEGER NOT NULL DEFAULT 0,
        cards_delta INTEGER NOT NULL DEFAULT 0,
        payload_json TEXT,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
      );

      CREATE TABLE imm_daily_rollups(
        rollup_day INTEGER NOT NULL,
        video_id INTEGER,
        total_sessions INTEGER NOT NULL DEFAULT 0,
        total_active_min REAL NOT NULL DEFAULT 0,
        total_lines_seen INTEGER NOT NULL DEFAULT 0,
        total_tokens_seen INTEGER NOT NULL DEFAULT 0,
        total_cards INTEGER NOT NULL DEFAULT 0,
        cards_per_hour REAL,
        tokens_per_min REAL,
        lookup_hit_rate REAL,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        PRIMARY KEY (rollup_day, video_id)
      );

      CREATE TABLE imm_monthly_rollups(
        rollup_month INTEGER NOT NULL,
        video_id INTEGER,
        total_sessions INTEGER NOT NULL DEFAULT 0,
        total_active_min REAL NOT NULL DEFAULT 0,
        total_lines_seen INTEGER NOT NULL DEFAULT 0,
        total_tokens_seen INTEGER NOT NULL DEFAULT 0,
        total_cards INTEGER NOT NULL DEFAULT 0,
        cards_per_hour REAL,
        tokens_per_min REAL,
        lookup_hit_rate REAL,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        PRIMARY KEY (rollup_month, video_id)
      );

      CREATE TABLE imm_words(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        headword TEXT NOT NULL,
        word TEXT NOT NULL,
        reading TEXT NOT NULL,
        part_of_speech TEXT,
        pos1 TEXT,
        pos2 TEXT,
        pos3 TEXT,
        first_seen INTEGER NOT NULL,
        last_seen INTEGER NOT NULL,
        frequency INTEGER NOT NULL DEFAULT 0,
        frequency_rank INTEGER,
        UNIQUE(headword, word, reading)
      );

      CREATE TABLE imm_kanji(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kanji TEXT NOT NULL UNIQUE,
        first_seen INTEGER NOT NULL,
        last_seen INTEGER NOT NULL,
        frequency INTEGER NOT NULL DEFAULT 0
      );

      CREATE TABLE imm_subtitle_lines(
        line_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        event_id INTEGER,
        video_id INTEGER NOT NULL,
        anime_id INTEGER,
        line_index INTEGER NOT NULL,
        segment_start_ms INTEGER,
        segment_end_ms INTEGER,
        text TEXT NOT NULL,
        secondary_text TEXT,
        CREATED_DATE INTEGER,
        LAST_UPDATE_DATE INTEGER,
        FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE,
        FOREIGN KEY(event_id) REFERENCES imm_session_events(event_id) ON DELETE SET NULL,
        FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE,
        FOREIGN KEY(anime_id) REFERENCES imm_anime(anime_id) ON DELETE SET NULL
      );

      CREATE TABLE imm_word_line_occurrences(
        line_id INTEGER NOT NULL,
        word_id INTEGER NOT NULL,
        occurrence_count INTEGER NOT NULL,
        PRIMARY KEY(line_id, word_id),
        FOREIGN KEY(line_id) REFERENCES imm_subtitle_lines(line_id) ON DELETE CASCADE,
        FOREIGN KEY(word_id) REFERENCES imm_words(id) ON DELETE CASCADE
      );

      CREATE TABLE imm_kanji_line_occurrences(
        line_id INTEGER NOT NULL,
        kanji_id INTEGER NOT NULL,
        occurrence_count INTEGER NOT NULL,
        PRIMARY KEY(line_id, kanji_id),
        FOREIGN KEY(line_id) REFERENCES imm_subtitle_lines(line_id) ON DELETE CASCADE,
        FOREIGN KEY(kanji_id) REFERENCES imm_kanji(id) ON DELETE CASCADE
      );

      CREATE TABLE imm_lifetime_global(
        global_id INTEGER PRIMARY KEY CHECK(global_id = 1),
        total_sessions INTEGER NOT NULL DEFAULT 0,
        total_active_ms INTEGER NOT NULL DEFAULT 0,
        total_cards INTEGER NOT NULL DEFAULT 0,
        active_days INTEGER NOT NULL DEFAULT 0,
        episodes_started INTEGER NOT NULL DEFAULT 0,
        episodes_completed INTEGER NOT NULL DEFAULT 0,
        anime_completed INTEGER NOT NULL DEFAULT 0,
        last_rebuilt_ms TEXT,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT
      );

      CREATE TABLE imm_lifetime_anime(
        anime_id INTEGER PRIMARY KEY,
        total_sessions INTEGER NOT NULL DEFAULT 0,
        total_active_ms INTEGER NOT NULL DEFAULT 0,
        total_cards INTEGER NOT NULL DEFAULT 0,
        total_lines_seen INTEGER NOT NULL DEFAULT 0,
        total_tokens_seen INTEGER NOT NULL DEFAULT 0,
        episodes_started INTEGER NOT NULL DEFAULT 0,
        episodes_completed INTEGER NOT NULL DEFAULT 0,
        first_watched_ms TEXT,
        last_watched_ms TEXT,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        FOREIGN KEY(anime_id) REFERENCES imm_anime(anime_id) ON DELETE CASCADE
      );

      CREATE TABLE imm_lifetime_media(
        video_id INTEGER PRIMARY KEY,
        total_sessions INTEGER NOT NULL DEFAULT 0,
        total_active_ms INTEGER NOT NULL DEFAULT 0,
        total_cards INTEGER NOT NULL DEFAULT 0,
        total_lines_seen INTEGER NOT NULL DEFAULT 0,
        total_tokens_seen INTEGER NOT NULL DEFAULT 0,
        completed INTEGER NOT NULL DEFAULT 0,
        first_watched_ms TEXT,
        last_watched_ms TEXT,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE
      );

      CREATE TABLE imm_lifetime_applied_sessions(
        session_id INTEGER PRIMARY KEY,
        applied_at_ms TEXT NOT NULL,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
      );

      CREATE TABLE imm_media_art(
        video_id INTEGER PRIMARY KEY,
        anilist_id INTEGER,
        cover_url TEXT,
        cover_blob BLOB,
        cover_blob_hash TEXT,
        fetched_at_ms TEXT,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE
      );

      CREATE TABLE imm_cover_art_blobs(
        blob_hash TEXT PRIMARY KEY,
        cover_blob BLOB NOT NULL,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT
      );

      CREATE TABLE imm_youtube_videos(
        video_id INTEGER PRIMARY KEY,
        youtube_video_id TEXT,
        video_url TEXT,
        video_title TEXT,
        video_thumbnail_url TEXT,
        channel_id TEXT,
        channel_name TEXT,
        channel_url TEXT,
        channel_thumbnail_url TEXT,
        uploader_id TEXT,
        uploader_url TEXT,
        description TEXT,
        metadata_json TEXT,
        fetched_at_ms TEXT NOT NULL,
        CREATED_DATE TEXT,
        LAST_UPDATE_DATE TEXT,
        FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE
      );

      INSERT INTO imm_videos (
        video_id, video_key, canonical_title, source_type, source_path, source_url, watched, duration_ms,
        CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'local:/tmp/repaired-event.mkv', 'Repaired Event', 1, '/tmp/repaired-event.mkv', NULL, 0, 0, '1000', '1000'
      );

      INSERT INTO imm_sessions (
        session_id, session_uuid, video_id, started_at_ms, status, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'session-1', 1, '1775940000000', 1, '1775940000000', '1775940000000'
      );

      INSERT INTO imm_session_events (
        event_id, session_id, ts_ms, event_type, line_index, segment_start_ms, segment_end_ms,
        tokens_delta, cards_delta, payload_json, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 1, -2147483648, 4, NULL, NULL, NULL, 0, 1, '{\"noteIds\":[1]}', '1775943304128', '1775943304128'
      );
    `);

    ensureSchema(db);

    const column = db.prepare(`PRAGMA table_info(imm_session_events)`).all() as Array<{
      name: string;
      type: string;
    }>;
    assert.equal(column.find((entry) => entry.name === 'ts_ms')?.type, 'TEXT');

    const row = db
      .prepare(
        `
        SELECT ts_ms AS tsMs, typeof(ts_ms) AS tsType, CREATED_DATE AS createdDate
        FROM imm_session_events
        WHERE event_id = 1
      `,
      )
      .get() as {
      tsMs: string;
      tsType: string;
      createdDate: string;
    };

    assert.equal(row.tsType, 'text');
    assert.equal(row.tsMs, '1775943304128');
    assert.equal(row.createdDate, '1775943304128');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('ensureSchema creates large-history performance indexes', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const indexNames = new Set(
      (
        db
          .prepare(`SELECT name FROM sqlite_master WHERE type = 'index' AND name LIKE 'idx_%'`)
          .all() as Array<{
          name: string;
        }>
      ).map((row) => row.name),
    );

    assert.ok(indexNames.has('idx_telemetry_sample_ms'));
    assert.ok(indexNames.has('idx_sessions_started_at'));
    assert.ok(indexNames.has('idx_sessions_ended_at'));
    assert.ok(indexNames.has('idx_words_frequency'));
    assert.ok(indexNames.has('idx_kanji_frequency'));
    assert.ok(indexNames.has('idx_media_art_anilist_id'));
    assert.ok(indexNames.has('idx_media_art_cover_url'));
    assert.ok(indexNames.has('idx_youtube_videos_channel_id'));
    assert.ok(indexNames.has('idx_youtube_videos_youtube_video_id'));
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('ensureSchema migrates legacy videos and backfills anime metadata from filenames', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    db.exec(`
      CREATE TABLE imm_schema_version (
        schema_version INTEGER PRIMARY KEY,
        applied_at_ms INTEGER NOT NULL
      );
      INSERT INTO imm_schema_version(schema_version, applied_at_ms) VALUES (4, 1);

      CREATE TABLE imm_videos(
        video_id INTEGER PRIMARY KEY AUTOINCREMENT,
        video_key TEXT NOT NULL UNIQUE,
        canonical_title TEXT NOT NULL,
        source_type INTEGER NOT NULL,
        source_path TEXT,
        source_url TEXT,
        duration_ms INTEGER NOT NULL CHECK(duration_ms>=0),
        file_size_bytes INTEGER CHECK(file_size_bytes>=0),
        codec_id INTEGER, container_id INTEGER,
        width_px INTEGER, height_px INTEGER, fps_x100 INTEGER,
        bitrate_kbps INTEGER, audio_codec_id INTEGER,
        hash_sha256 TEXT, screenshot_path TEXT,
        metadata_json TEXT,
        CREATED_DATE INTEGER,
        LAST_UPDATE_DATE INTEGER
      );
    `);

    const insertLegacyVideo = db.prepare(`
      INSERT INTO imm_videos (
        video_key, canonical_title, source_type, source_path, source_url,
        duration_ms, file_size_bytes, codec_id, container_id, width_px, height_px,
        fps_x100, bitrate_kbps, audio_codec_id, hash_sha256, screenshot_path,
        metadata_json, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    insertLegacyVideo.run(
      'local:/library/Little Witch Academia S02E05.mkv',
      'Episode 5',
      SOURCE_TYPE_LOCAL,
      '/library/Little Witch Academia S02E05.mkv',
      null,
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
      1,
      1,
    );
    insertLegacyVideo.run(
      'local:/library/Little Witch Academia S02E06.mkv',
      'Episode 6',
      SOURCE_TYPE_LOCAL,
      '/library/Little Witch Academia S02E06.mkv',
      null,
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
      1,
      1,
    );
    insertLegacyVideo.run(
      'local:/library/[SubsPlease] Frieren - 03 - Departure.mkv',
      'Episode 3',
      SOURCE_TYPE_LOCAL,
      '/library/[SubsPlease] Frieren - 03 - Departure.mkv',
      null,
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
      1,
      1,
    );

    ensureSchema(db);

    const videoColumns = new Set(
      (
        db.prepare('PRAGMA table_info(imm_videos)').all() as Array<{
          name: string;
        }>
      ).map((row) => row.name),
    );
    assert.ok(videoColumns.has('anime_id'));
    assert.ok(videoColumns.has('parsed_basename'));
    assert.ok(videoColumns.has('parsed_title'));
    assert.ok(videoColumns.has('parsed_season'));
    assert.ok(videoColumns.has('parsed_episode'));
    assert.ok(videoColumns.has('parser_source'));
    assert.ok(videoColumns.has('parser_confidence'));
    assert.ok(videoColumns.has('parse_metadata_json'));

    const animeRows = db
      .prepare('SELECT canonical_title FROM imm_anime ORDER BY canonical_title')
      .all() as Array<{ canonical_title: string }>;
    assert.deepEqual(
      animeRows.map((row) => row.canonical_title),
      ['Frieren', 'Little Witch Academia Season 2'],
    );

    const littleWitchRows = db
      .prepare(
        `
          SELECT
            a.canonical_title AS anime_title,
            v.parsed_title,
            v.parsed_basename,
            v.parsed_season,
            v.parsed_episode,
            v.parser_source,
            v.parser_confidence
          FROM imm_videos v
          JOIN imm_anime a ON a.anime_id = v.anime_id
          WHERE v.video_key LIKE 'local:/library/Little Witch Academia%'
          ORDER BY v.video_key
        `,
      )
      .all() as Array<{
      anime_title: string;
      parsed_title: string | null;
      parsed_basename: string | null;
      parsed_season: number | null;
      parsed_episode: number | null;
      parser_source: string | null;
      parser_confidence: number | null;
    }>;

    assert.equal(littleWitchRows.length, 2);
    assert.deepEqual(
      littleWitchRows.map((row) => ({
        animeTitle: row.anime_title,
        parsedTitle: row.parsed_title,
        parsedBasename: row.parsed_basename,
        parsedSeason: row.parsed_season,
        parsedEpisode: row.parsed_episode,
        parserSource: row.parser_source,
      })),
      [
        {
          animeTitle: 'Little Witch Academia Season 2',
          parsedTitle: 'Little Witch Academia',
          parsedBasename: 'Little Witch Academia S02E05.mkv',
          parsedSeason: 2,
          parsedEpisode: 5,
          parserSource: 'fallback',
        },
        {
          animeTitle: 'Little Witch Academia Season 2',
          parsedTitle: 'Little Witch Academia',
          parsedBasename: 'Little Witch Academia S02E06.mkv',
          parsedSeason: 2,
          parsedEpisode: 6,
          parserSource: 'fallback',
        },
      ],
    );
    assert.ok(
      littleWitchRows.every(
        (row) => typeof row.parser_confidence === 'number' && row.parser_confidence > 0,
      ),
    );

    const frierenRow = db
      .prepare(
        `
          SELECT
            a.canonical_title AS anime_title,
            v.parsed_title,
            v.parsed_episode,
            v.parser_source
          FROM imm_videos v
          JOIN imm_anime a ON a.anime_id = v.anime_id
          WHERE v.video_key = ?
        `,
      )
      .get('local:/library/[SubsPlease] Frieren - 03 - Departure.mkv') as {
      anime_title: string;
      parsed_title: string | null;
      parsed_episode: number | null;
      parser_source: string | null;
    } | null;

    assert.ok(frierenRow);
    assert.equal(frierenRow?.anime_title, 'Frieren');
    assert.equal(frierenRow?.parsed_title, 'Frieren');
    assert.equal(frierenRow?.parsed_episode, 3);
    assert.equal(frierenRow?.parser_source, 'fallback');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('ensureSchema adds subtitle-line occurrence tables to schema version 6 databases', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    db.exec(`
      CREATE TABLE imm_schema_version (
        schema_version INTEGER PRIMARY KEY,
        applied_at_ms INTEGER NOT NULL
      );
      INSERT INTO imm_schema_version(schema_version, applied_at_ms) VALUES (6, 1);

      CREATE TABLE imm_videos(
        video_id INTEGER PRIMARY KEY AUTOINCREMENT,
        video_key TEXT NOT NULL UNIQUE,
        anime_id INTEGER,
        canonical_title TEXT NOT NULL,
        source_type INTEGER NOT NULL,
        source_path TEXT,
        source_url TEXT,
        parsed_basename TEXT,
        parsed_title TEXT,
        parsed_season INTEGER,
        parsed_episode INTEGER,
        parser_source TEXT,
        parser_confidence REAL,
        parse_metadata_json TEXT,
        duration_ms INTEGER NOT NULL CHECK(duration_ms>=0),
        file_size_bytes INTEGER CHECK(file_size_bytes>=0),
        codec_id INTEGER, container_id INTEGER,
        width_px INTEGER, height_px INTEGER, fps_x100 INTEGER,
        bitrate_kbps INTEGER, audio_codec_id INTEGER,
        hash_sha256 TEXT, screenshot_path TEXT,
        metadata_json TEXT,
        CREATED_DATE INTEGER,
        LAST_UPDATE_DATE INTEGER
      );
      CREATE TABLE imm_sessions(
        session_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_uuid TEXT NOT NULL UNIQUE,
        video_id INTEGER NOT NULL,
        started_at_ms INTEGER NOT NULL,
        ended_at_ms INTEGER,
        status INTEGER NOT NULL,
        locale_id INTEGER,
        target_lang_id INTEGER,
        difficulty_tier INTEGER,
        subtitle_mode INTEGER,
        CREATED_DATE INTEGER,
        LAST_UPDATE_DATE INTEGER
      );
      CREATE TABLE imm_session_events(
        event_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        ts_ms INTEGER NOT NULL,
        event_type INTEGER NOT NULL,
        line_index INTEGER,
        segment_start_ms INTEGER,
        segment_end_ms INTEGER,
        words_delta INTEGER NOT NULL DEFAULT 0,
        cards_delta INTEGER NOT NULL DEFAULT 0,
        payload_json TEXT,
        CREATED_DATE INTEGER,
        LAST_UPDATE_DATE INTEGER
      );
      CREATE TABLE imm_words(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        headword TEXT,
        word TEXT,
        reading TEXT,
        part_of_speech TEXT,
        pos1 TEXT,
        pos2 TEXT,
        pos3 TEXT,
        first_seen REAL,
        last_seen REAL,
        frequency INTEGER,
        UNIQUE(headword, word, reading)
      );
      CREATE TABLE imm_kanji(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kanji TEXT,
        first_seen REAL,
        last_seen REAL,
        frequency INTEGER,
        UNIQUE(kanji)
      );
      CREATE TABLE imm_rollup_state(
        state_key TEXT PRIMARY KEY,
        state_value INTEGER NOT NULL
      );
    `);

    ensureSchema(db);

    const tableNames = new Set(
      (
        db
          .prepare(`SELECT name FROM sqlite_master WHERE type = 'table' AND name LIKE 'imm_%'`)
          .all() as Array<{ name: string }>
      ).map((row) => row.name),
    );

    assert.ok(tableNames.has('imm_subtitle_lines'));
    assert.ok(tableNames.has('imm_word_line_occurrences'));
    assert.ok(tableNames.has('imm_kanji_line_occurrences'));
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('ensureSchema migrates legacy cover art blobs into the shared blob store', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    db.prepare('UPDATE imm_schema_version SET schema_version = 12').run();

    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/legacy-cover-art.mkv', {
      canonicalTitle: 'Legacy Cover Art',
      sourcePath: '/tmp/legacy-cover-art.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const legacyBlob = Uint8Array.from([0xde, 0xad, 0xbe, 0xef]);

    db.prepare(
      `
        INSERT INTO imm_media_art (
          video_id,
          anilist_id,
          cover_url,
          cover_blob,
          cover_blob_hash,
          title_romaji,
          title_english,
          episodes_total,
          fetched_at_ms,
          CREATED_DATE,
          LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    ).run(videoId, null, null, legacyBlob, null, null, null, null, 1, 1, 1);

    assert.doesNotThrow(() => ensureSchema(db));

    const mediaArtRow = db
      .prepare(
        'SELECT cover_blob AS coverBlob, cover_blob_hash AS coverBlobHash FROM imm_media_art',
      )
      .get() as {
      coverBlob: ArrayBuffer | Uint8Array | Buffer | null;
      coverBlobHash: string | null;
    } | null;

    assert.ok(mediaArtRow);
    assert.ok(mediaArtRow?.coverBlobHash);
    assert.equal(
      parseCoverBlobReference(normalizeCoverBlobBytes(mediaArtRow?.coverBlob)),
      mediaArtRow?.coverBlobHash,
    );

    const sharedBlobRow = db
      .prepare('SELECT cover_blob AS coverBlob FROM imm_cover_art_blobs WHERE blob_hash = ?')
      .get(mediaArtRow?.coverBlobHash) as {
      coverBlob: ArrayBuffer | Uint8Array | Buffer;
    } | null;

    assert.ok(sharedBlobRow);
    assert.equal(normalizeCoverBlobBytes(sharedBlobRow?.coverBlob)?.toString('hex'), 'deadbeef');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('anime rows are reused by normalized parsed title and upgraded with AniList metadata', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const firstVideoId = getOrCreateVideoRecord(db, 'local:/tmp/lwa-s02e05.mkv', {
      canonicalTitle: 'Episode 5',
      sourcePath: '/tmp/Little Witch Academia S02E05.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const secondVideoId = getOrCreateVideoRecord(db, 'local:/tmp/lwa-s02e06.mkv', {
      canonicalTitle: 'Episode 6',
      sourcePath: '/tmp/Little Witch Academia S02E06.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });

    const provisionalAnimeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Little Witch Academia',
      canonicalTitle: 'Little Witch Academia',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: '{"source":"parsed"}',
    });
    linkVideoToAnimeRecord(db, firstVideoId, {
      animeId: provisionalAnimeId,
      parsedBasename: 'Little Witch Academia S02E05.mkv',
      parsedTitle: 'Little Witch Academia',
      parsedSeason: 2,
      parsedEpisode: 5,
      parserSource: 'fallback',
      parserConfidence: 0.6,
      parseMetadataJson: '{"source":"parsed","episode":5}',
    });

    const reusedAnimeId = getOrCreateAnimeRecord(db, {
      parsedTitle: ' little  witch academia ',
      canonicalTitle: 'Little Witch Academia',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: '{"source":"parsed"}',
    });
    linkVideoToAnimeRecord(db, secondVideoId, {
      animeId: reusedAnimeId,
      parsedBasename: 'Little Witch Academia S02E06.mkv',
      parsedTitle: 'Little Witch Academia',
      parsedSeason: 2,
      parsedEpisode: 6,
      parserSource: 'fallback',
      parserConfidence: 0.6,
      parseMetadataJson: '{"source":"parsed","episode":6}',
    });

    assert.equal(reusedAnimeId, provisionalAnimeId);

    const upgradedAnimeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Little Witch Academia',
      canonicalTitle: 'Little Witch Academia TV',
      anilistId: 33_435,
      titleRomaji: 'Little Witch Academia',
      titleEnglish: 'Little Witch Academia',
      titleNative: 'リトルウィッチアカデミア',
      metadataJson: '{"source":"anilist"}',
    });

    assert.equal(upgradedAnimeId, provisionalAnimeId);

    const animeRows = db.prepare('SELECT * FROM imm_anime').all() as Array<{
      anime_id: number;
      normalized_title_key: string;
      canonical_title: string;
      anilist_id: number | null;
      title_romaji: string | null;
      title_english: string | null;
      title_native: string | null;
      metadata_json: string | null;
    }>;
    assert.equal(animeRows.length, 1);
    assert.equal(animeRows[0]?.anime_id, provisionalAnimeId);
    assert.equal(animeRows[0]?.normalized_title_key, 'little witch academia');
    assert.equal(animeRows[0]?.canonical_title, 'Little Witch Academia TV');
    assert.equal(animeRows[0]?.anilist_id, 33_435);
    assert.equal(animeRows[0]?.title_romaji, 'Little Witch Academia');
    assert.equal(animeRows[0]?.title_english, 'Little Witch Academia');
    assert.equal(animeRows[0]?.title_native, 'リトルウィッチアカデミア');
    assert.equal(animeRows[0]?.metadata_json, '{"source":"anilist"}');

    const linkedVideos = db
      .prepare(
        `
          SELECT anime_id, parsed_title, parsed_season, parsed_episode
          FROM imm_videos
          WHERE video_id IN (?, ?)
          ORDER BY video_id
        `,
      )
      .all(firstVideoId, secondVideoId) as Array<{
      anime_id: number | null;
      parsed_title: string | null;
      parsed_season: number | null;
      parsed_episode: number | null;
    }>;

    assert.deepEqual(linkedVideos, [
      {
        anime_id: provisionalAnimeId,
        parsed_title: 'Little Witch Academia',
        parsed_season: 2,
        parsed_episode: 5,
      },
      {
        anime_id: provisionalAnimeId,
        parsed_title: 'Little Witch Academia',
        parsed_season: 2,
        parsed_episode: 6,
      },
    ]);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('youtube videos can be regrouped under a shared channel anime identity', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);

    const firstVideoId = getOrCreateVideoRecord(
      db,
      'remote:https://www.youtube.com/watch?v=video-1',
      {
        canonicalTitle: 'watch?v video-1',
        sourcePath: null,
        sourceUrl: 'https://www.youtube.com/watch?v=video-1',
        sourceType: SOURCE_TYPE_REMOTE,
      },
    );
    const secondVideoId = getOrCreateVideoRecord(
      db,
      'remote:https://www.youtube.com/watch?v=video-2',
      {
        canonicalTitle: 'watch?v video-2',
        sourcePath: null,
        sourceUrl: 'https://www.youtube.com/watch?v=video-2',
        sourceType: SOURCE_TYPE_REMOTE,
      },
    );

    const firstAnimeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'watch?v video-1',
      canonicalTitle: 'watch?v video-1',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    linkVideoToAnimeRecord(db, firstVideoId, {
      animeId: firstAnimeId,
      parsedBasename: null,
      parsedTitle: 'watch?v video-1',
      parsedSeason: null,
      parsedEpisode: null,
      parserSource: 'fallback',
      parserConfidence: 0.2,
      parseMetadataJson: '{"source":"fallback"}',
    });

    const secondAnimeId = getOrCreateAnimeRecord(db, {
      parsedTitle: 'watch?v video-2',
      canonicalTitle: 'watch?v video-2',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    linkVideoToAnimeRecord(db, secondVideoId, {
      animeId: secondAnimeId,
      parsedBasename: null,
      parsedTitle: 'watch?v video-2',
      parsedSeason: null,
      parsedEpisode: null,
      parserSource: 'fallback',
      parserConfidence: 0.2,
      parseMetadataJson: '{"source":"fallback"}',
    });

    linkYoutubeVideoToAnimeRecord(db, firstVideoId, {
      youtubeVideoId: 'video-1',
      videoUrl: 'https://www.youtube.com/watch?v=video-1',
      videoTitle: 'Video One',
      videoThumbnailUrl: 'https://i.ytimg.com/vi/video-1/hqdefault.jpg',
      channelId: 'UC123',
      channelName: 'Channel Name',
      channelUrl: 'https://www.youtube.com/channel/UC123',
      channelThumbnailUrl:
        'https://yt3.googleusercontent.com/channel-123=s176-c-k-c0x00ffffff-no-rj',
      uploaderId: '@channelname',
      uploaderUrl: 'https://www.youtube.com/@channelname',
      description: null,
      metadataJson: '{"id":"video-1"}',
    });
    linkYoutubeVideoToAnimeRecord(db, secondVideoId, {
      youtubeVideoId: 'video-2',
      videoUrl: 'https://www.youtube.com/watch?v=video-2',
      videoTitle: 'Video Two',
      videoThumbnailUrl: 'https://i.ytimg.com/vi/video-2/hqdefault.jpg',
      channelId: 'UC123',
      channelName: 'Channel Name',
      channelUrl: 'https://www.youtube.com/channel/UC123',
      channelThumbnailUrl:
        'https://yt3.googleusercontent.com/channel-123=s176-c-k-c0x00ffffff-no-rj',
      uploaderId: '@channelname',
      uploaderUrl: 'https://www.youtube.com/@channelname',
      description: null,
      metadataJson: '{"id":"video-2"}',
    });

    const animeRows = db.prepare('SELECT anime_id, canonical_title FROM imm_anime').all() as Array<{
      anime_id: number;
      canonical_title: string;
    }>;
    const videoRows = db
      .prepare('SELECT video_id, anime_id, parsed_title FROM imm_videos ORDER BY video_id ASC')
      .all() as Array<{ video_id: number; anime_id: number | null; parsed_title: string | null }>;

    const channelAnimeRows = animeRows.filter((row) => row.canonical_title === 'Channel Name');
    assert.equal(channelAnimeRows.length, 1);
    assert.equal(videoRows[0]?.anime_id, channelAnimeRows[0]?.anime_id);
    assert.equal(videoRows[1]?.anime_id, channelAnimeRows[0]?.anime_id);
    assert.equal(videoRows[0]?.parsed_title, 'Channel Name');
    assert.equal(videoRows[1]?.parsed_title, 'Channel Name');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('start/finalize session updates ended_at and status', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/slice-a.mkv', {
      canonicalTitle: 'Slice A Episode',
      sourcePath: '/tmp/slice-a.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const startedAtMs = 1_234_567_000;
    const endedAtMs = startedAtMs + 8_500;
    const { sessionId, state } = startSessionRecord(db, videoId, startedAtMs);

    finalizeSessionRecord(db, state, endedAtMs);

    const row = db
      .prepare('SELECT ended_at_ms, status FROM imm_sessions WHERE session_id = ?')
      .get(sessionId) as {
      ended_at_ms: string | null;
      status: number;
    } | null;

    assert.ok(row);
    assert.equal(Number(row?.ended_at_ms ?? 0), endedAtMs);
    assert.equal(row?.status, SESSION_STATUS_ENDED);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('finalize session persists ended media position', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/slice-a-ended-media.mkv', {
      canonicalTitle: 'Slice A Ended Media',
      sourcePath: '/tmp/slice-a-ended-media.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const startedAtMs = 1_234_567_000;
    const endedAtMs = startedAtMs + 8_500;
    const { sessionId, state } = startSessionRecord(db, videoId, startedAtMs);
    state.lastMediaMs = 91_000;

    finalizeSessionRecord(db, state, endedAtMs);

    const row = db
      .prepare('SELECT ended_media_ms FROM imm_sessions WHERE session_id = ?')
      .get(sessionId) as {
      ended_media_ms: number | null;
    } | null;

    assert.ok(row);
    assert.equal(row?.ended_media_ms, 91_000);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('executeQueuedWrite inserts event and telemetry rows', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/slice-a-events.mkv', {
      canonicalTitle: 'Slice A Events',
      sourcePath: '/tmp/slice-a-events.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const { sessionId } = startSessionRecord(db, videoId, 5_000);

    executeQueuedWrite(
      {
        kind: 'telemetry',
        sessionId,
        sampleMs: 6_000,
        totalWatchedMs: 1_000,
        activeWatchedMs: 900,
        linesSeen: 3,
        tokensSeen: 6,
        cardsMined: 1,
        lookupCount: 2,
        lookupHits: 1,
        yomitanLookupCount: 0,
        pauseCount: 1,
        pauseMs: 50,
        seekForwardCount: 0,
        seekBackwardCount: 0,
        mediaBufferEvents: 0,
      },
      stmts,
    );
    executeQueuedWrite(
      {
        kind: 'event',
        sessionId,
        sampleMs: 6_100,
        eventType: EVENT_SUBTITLE_LINE,
        lineIndex: 1,
        segmentStartMs: 0,
        segmentEndMs: 800,
        tokensDelta: 2,
        cardsDelta: 0,
        payloadJson: '{"event":"subtitle-line"}',
      },
      stmts,
    );

    const telemetryCount = db
      .prepare('SELECT COUNT(*) AS total FROM imm_session_telemetry WHERE session_id = ?')
      .get(sessionId) as { total: number };
    const eventCount = db
      .prepare('SELECT COUNT(*) AS total FROM imm_session_events WHERE session_id = ?')
      .get(sessionId) as { total: number };

    assert.equal(telemetryCount.total, 1);
    assert.equal(eventCount.total, 1);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('executeQueuedWrite rejects partial telemetry writes instead of zero-filling', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/partial-telemetry.mkv', {
      canonicalTitle: 'Partial Telemetry',
      sourcePath: '/tmp/partial-telemetry.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const { sessionId } = startSessionRecord(db, videoId, 5_000);

    assert.throws(
      () =>
        executeQueuedWrite(
          {
            kind: 'telemetry',
            sessionId,
            sampleMs: 6_000,
            totalWatchedMs: 1_000,
            activeWatchedMs: 900,
            linesSeen: 3,
            cardsMined: 1,
            lookupCount: 2,
            lookupHits: 1,
            yomitanLookupCount: 0,
            pauseCount: 1,
            pauseMs: 50,
            seekForwardCount: 0,
            seekBackwardCount: 0,
            mediaBufferEvents: 0,
          },
          stmts,
        ),
      /Incomplete telemetry write/,
    );

    const telemetryCount = db
      .prepare('SELECT COUNT(*) AS total FROM imm_session_telemetry WHERE session_id = ?')
      .get(sessionId) as { total: number };
    assert.equal(telemetryCount.total, 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('executeQueuedWrite inserts and upserts word and kanji rows', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    stmts.wordUpsertStmt.run('猫', '猫', '', 'noun', '名詞', '一般', '', 10.0, 10.0);
    stmts.wordUpsertStmt.run('猫', '猫', '', 'noun', '名詞', '一般', '', 5.0, 15.0);
    stmts.kanjiUpsertStmt.run('日', 9.0, 9.0);
    stmts.kanjiUpsertStmt.run('日', 8.0, 11.0);

    const wordRow = db
      .prepare(
        `SELECT headword, frequency, part_of_speech, pos1, pos2, first_seen, last_seen
         FROM imm_words WHERE headword = ?`,
      )
      .get('猫') as {
      headword: string;
      frequency: number;
      part_of_speech: string;
      pos1: string;
      pos2: string;
      first_seen: number;
      last_seen: number;
    } | null;
    const kanjiRow = db
      .prepare('SELECT kanji, frequency, first_seen, last_seen FROM imm_kanji WHERE kanji = ?')
      .get('日') as {
      kanji: string;
      frequency: number;
      first_seen: number;
      last_seen: number;
    } | null;

    assert.ok(wordRow);
    assert.ok(kanjiRow);
    assert.equal(wordRow?.frequency, 2);
    assert.equal(wordRow?.part_of_speech, 'noun');
    assert.equal(wordRow?.pos1, '名詞');
    assert.equal(wordRow?.pos2, '一般');
    assert.equal(kanjiRow?.frequency, 2);
    assert.equal(wordRow?.first_seen, 5);
    assert.equal(wordRow?.last_seen, 15);
    assert.equal(kanjiRow?.first_seen, 8);
    assert.equal(kanjiRow?.last_seen, 11);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('word upsert replaces legacy other part_of_speech when better POS metadata arrives later', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    stmts.wordUpsertStmt.run(
      '知っている',
      '知っている',
      'しっている',
      'other',
      '動詞',
      '自立',
      '',
      10,
      10,
    );
    stmts.wordUpsertStmt.run(
      '知っている',
      '知っている',
      'しっている',
      'verb',
      '動詞',
      '自立',
      '',
      11,
      12,
    );

    const row = db
      .prepare('SELECT frequency, part_of_speech, pos1, pos2 FROM imm_words WHERE headword = ?')
      .get('知っている') as {
      frequency: number;
      part_of_speech: string;
      pos1: string;
      pos2: string;
    } | null;

    assert.ok(row);
    assert.equal(row?.frequency, 2);
    assert.equal(row?.part_of_speech, 'verb');
    assert.equal(row?.pos1, '動詞');
    assert.equal(row?.pos2, '自立');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});
