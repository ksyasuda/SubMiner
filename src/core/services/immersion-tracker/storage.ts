import { parseMediaInfo } from '../../../jimaku/utils';
import type { DatabaseSync } from './sqlite';
import { SCHEMA_VERSION } from './types';
import type { QueuedWrite, VideoMetadata } from './types';

export interface TrackerPreparedStatements {
  telemetryInsertStmt: ReturnType<DatabaseSync['prepare']>;
  eventInsertStmt: ReturnType<DatabaseSync['prepare']>;
  wordUpsertStmt: ReturnType<DatabaseSync['prepare']>;
  kanjiUpsertStmt: ReturnType<DatabaseSync['prepare']>;
  subtitleLineInsertStmt: ReturnType<DatabaseSync['prepare']>;
  wordIdSelectStmt: ReturnType<DatabaseSync['prepare']>;
  kanjiIdSelectStmt: ReturnType<DatabaseSync['prepare']>;
  wordLineOccurrenceUpsertStmt: ReturnType<DatabaseSync['prepare']>;
  kanjiLineOccurrenceUpsertStmt: ReturnType<DatabaseSync['prepare']>;
  videoAnimeIdSelectStmt: ReturnType<DatabaseSync['prepare']>;
}

export interface AnimeRecordInput {
  parsedTitle: string;
  canonicalTitle: string;
  anilistId: number | null;
  titleRomaji: string | null;
  titleEnglish: string | null;
  titleNative: string | null;
  metadataJson: string | null;
}

export interface VideoAnimeLinkInput {
  animeId: number | null;
  parsedBasename: string | null;
  parsedTitle: string | null;
  parsedSeason: number | null;
  parsedEpisode: number | null;
  parserSource: string | null;
  parserConfidence: number | null;
  parseMetadataJson: string | null;
}

function hasColumn(db: DatabaseSync, tableName: string, columnName: string): boolean {
  return db
    .prepare(`PRAGMA table_info(${tableName})`)
    .all()
    .some((row: unknown) => (row as { name: string }).name === columnName);
}

function addColumnIfMissing(
  db: DatabaseSync,
  tableName: string,
  columnName: string,
  columnType = 'INTEGER',
): void {
  if (!hasColumn(db, tableName, columnName)) {
    db.exec(`ALTER TABLE ${tableName} ADD COLUMN ${columnName} ${columnType}`);
  }
}

function dropColumnIfExists(db: DatabaseSync, tableName: string, columnName: string): void {
  if (hasColumn(db, tableName, columnName)) {
    db.exec(`ALTER TABLE ${tableName} DROP COLUMN ${columnName}`);
  }
}

export function applyPragmas(db: DatabaseSync): void {
  db.exec('PRAGMA journal_mode = WAL');
  db.exec('PRAGMA synchronous = NORMAL');
  db.exec('PRAGMA foreign_keys = ON');
  db.exec('PRAGMA busy_timeout = 2500');
}

export function normalizeAnimeIdentityKey(title: string): string {
  return title
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function looksLikeEpisodeOnlyTitle(title: string): boolean {
  const normalized = title
    .normalize('NFKC')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
  return /^(episode|ep)\s*\d{1,3}$/.test(normalized) || /^第\s*\d{1,3}\s*話$/.test(normalized);
}

function parserConfidenceToScore(confidence: 'high' | 'medium' | 'low'): number {
  switch (confidence) {
    case 'high':
      return 1;
    case 'medium':
      return 0.6;
    default:
      return 0.2;
  }
}

function parseLegacyAnimeBackfillCandidate(
  sourcePath: string | null,
  canonicalTitle: string,
): {
  basename: string | null;
  title: string;
  season: number | null;
  episode: number | null;
  source: 'fallback';
  confidenceScore: number;
  metadataJson: string;
} | null {
  const fromPath =
    sourcePath && sourcePath.trim().length > 0 ? parseMediaInfo(sourcePath.trim()) : null;
  if (fromPath?.title && !looksLikeEpisodeOnlyTitle(fromPath.title)) {
    return {
      basename: fromPath.filename || null,
      title: fromPath.title,
      season: fromPath.season,
      episode: fromPath.episode,
      source: 'fallback',
      confidenceScore: parserConfidenceToScore(fromPath.confidence),
      metadataJson: JSON.stringify({
        confidence: fromPath.confidence,
        filename: fromPath.filename,
        rawTitle: fromPath.rawTitle,
        migrationSource: 'source_path',
      }),
    };
  }

  const fallbackTitle = canonicalTitle.trim();
  if (!fallbackTitle) return null;
  const fromTitle = parseMediaInfo(fallbackTitle);
  if (!fromTitle.title || looksLikeEpisodeOnlyTitle(fromTitle.title)) {
    return null;
  }

  return {
    basename: null,
    title: fromTitle.title,
    season: fromTitle.season,
    episode: fromTitle.episode,
    source: 'fallback',
    confidenceScore: parserConfidenceToScore(fromTitle.confidence),
    metadataJson: JSON.stringify({
      confidence: fromTitle.confidence,
      filename: fromTitle.filename,
      rawTitle: fromTitle.rawTitle,
      migrationSource: 'canonical_title',
    }),
  };
}

export function getOrCreateAnimeRecord(db: DatabaseSync, input: AnimeRecordInput): number {
  const normalizedTitleKey = normalizeAnimeIdentityKey(input.parsedTitle);
  if (!normalizedTitleKey) {
    throw new Error('parsedTitle is required to create or update an anime record');
  }

  const byAnilistId =
    input.anilistId !== null
      ? (db.prepare('SELECT anime_id FROM imm_anime WHERE anilist_id = ?').get(input.anilistId) as {
          anime_id: number;
        } | null)
      : null;
  const byNormalizedTitle = db
    .prepare('SELECT anime_id FROM imm_anime WHERE normalized_title_key = ?')
    .get(normalizedTitleKey) as { anime_id: number } | null;
  const existing = byAnilistId ?? byNormalizedTitle;
  if (existing?.anime_id) {
    db.prepare(
      `
        UPDATE imm_anime
        SET
          canonical_title = COALESCE(NULLIF(?, ''), canonical_title),
          anilist_id = COALESCE(?, anilist_id),
          title_romaji = COALESCE(?, title_romaji),
          title_english = COALESCE(?, title_english),
          title_native = COALESCE(?, title_native),
          metadata_json = COALESCE(?, metadata_json),
          LAST_UPDATE_DATE = ?
        WHERE anime_id = ?
      `,
    ).run(
      input.canonicalTitle,
      input.anilistId,
      input.titleRomaji,
      input.titleEnglish,
      input.titleNative,
      input.metadataJson,
      Date.now(),
      existing.anime_id,
    );
    return existing.anime_id;
  }

  const nowMs = Date.now();
  const result = db
    .prepare(
      `
        INSERT INTO imm_anime(
          normalized_title_key,
          canonical_title,
          anilist_id,
          title_romaji,
          title_english,
          title_native,
          metadata_json,
          CREATED_DATE,
          LAST_UPDATE_DATE
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
    )
    .run(
      normalizedTitleKey,
      input.canonicalTitle,
      input.anilistId,
      input.titleRomaji,
      input.titleEnglish,
      input.titleNative,
      input.metadataJson,
      nowMs,
      nowMs,
    );
  return Number(result.lastInsertRowid);
}

export function linkVideoToAnimeRecord(
  db: DatabaseSync,
  videoId: number,
  input: VideoAnimeLinkInput,
): void {
  db.prepare(
    `
      UPDATE imm_videos
      SET
        anime_id = ?,
        parsed_basename = ?,
        parsed_title = ?,
        parsed_season = ?,
        parsed_episode = ?,
        parser_source = ?,
        parser_confidence = ?,
        parse_metadata_json = ?,
        LAST_UPDATE_DATE = ?
      WHERE video_id = ?
    `,
  ).run(
    input.animeId,
    input.parsedBasename,
    input.parsedTitle,
    input.parsedSeason,
    input.parsedEpisode,
    input.parserSource,
    input.parserConfidence,
    input.parseMetadataJson,
    Date.now(),
    videoId,
  );
}

function migrateLegacyAnimeMetadata(db: DatabaseSync): void {
  addColumnIfMissing(db, 'imm_videos', 'anime_id', 'INTEGER REFERENCES imm_anime(anime_id)');
  addColumnIfMissing(db, 'imm_videos', 'parsed_basename', 'TEXT');
  addColumnIfMissing(db, 'imm_videos', 'parsed_title', 'TEXT');
  addColumnIfMissing(db, 'imm_videos', 'parsed_season', 'INTEGER');
  addColumnIfMissing(db, 'imm_videos', 'parsed_episode', 'INTEGER');
  addColumnIfMissing(db, 'imm_videos', 'parser_source', 'TEXT');
  addColumnIfMissing(db, 'imm_videos', 'parser_confidence', 'REAL');
  addColumnIfMissing(db, 'imm_videos', 'parse_metadata_json', 'TEXT');

  const legacyRows = db
    .prepare(
      `
        SELECT video_id, source_path, canonical_title
        FROM imm_videos
        WHERE anime_id IS NULL
      `,
    )
    .all() as Array<{
    video_id: number;
    source_path: string | null;
    canonical_title: string;
  }>;

  for (const row of legacyRows) {
    const parsed = parseLegacyAnimeBackfillCandidate(row.source_path, row.canonical_title);
    if (!parsed) continue;

    const animeId = getOrCreateAnimeRecord(db, {
      parsedTitle: parsed.title,
      canonicalTitle: parsed.title,
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: parsed.metadataJson,
    });
    linkVideoToAnimeRecord(db, row.video_id, {
      animeId,
      parsedBasename: parsed.basename,
      parsedTitle: parsed.title,
      parsedSeason: parsed.season,
      parsedEpisode: parsed.episode,
      parserSource: parsed.source,
      parserConfidence: parsed.confidenceScore,
      parseMetadataJson: parsed.metadataJson,
    });
  }
}

export function ensureSchema(db: DatabaseSync): void {
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_schema_version (
      schema_version INTEGER PRIMARY KEY,
      applied_at_ms INTEGER NOT NULL
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_rollup_state(
      state_key TEXT PRIMARY KEY,
      state_value INTEGER NOT NULL
    );
  `);
  db.exec(`
    INSERT INTO imm_rollup_state(state_key, state_value)
    VALUES ('last_rollup_sample_ms', 0)
    ON CONFLICT(state_key) DO NOTHING
  `);

  const currentVersion = db
    .prepare('SELECT schema_version FROM imm_schema_version ORDER BY schema_version DESC LIMIT 1')
    .get() as { schema_version: number } | null;
  if (currentVersion?.schema_version === SCHEMA_VERSION) {
    return;
  }

  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_anime(
      anime_id INTEGER PRIMARY KEY AUTOINCREMENT,
      normalized_title_key TEXT NOT NULL UNIQUE,
      canonical_title TEXT NOT NULL,
      anilist_id INTEGER UNIQUE,
      title_romaji TEXT,
      title_english TEXT,
      title_native TEXT,
      episodes_total INTEGER,
      metadata_json TEXT,
      CREATED_DATE INTEGER,
      LAST_UPDATE_DATE INTEGER
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_videos(
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
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_sessions(
      session_id INTEGER PRIMARY KEY AUTOINCREMENT,
      session_uuid TEXT NOT NULL UNIQUE,
      video_id INTEGER NOT NULL,
      started_at_ms INTEGER NOT NULL, ended_at_ms INTEGER,
      status INTEGER NOT NULL,
      locale_id INTEGER, target_lang_id INTEGER,
      difficulty_tier INTEGER, subtitle_mode INTEGER,
      CREATED_DATE INTEGER,
      LAST_UPDATE_DATE INTEGER,
      FOREIGN KEY(video_id) REFERENCES imm_videos(video_id)
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_session_telemetry(
      telemetry_id INTEGER PRIMARY KEY AUTOINCREMENT,
      session_id INTEGER NOT NULL,
      sample_ms INTEGER NOT NULL,
      total_watched_ms INTEGER NOT NULL DEFAULT 0,
      active_watched_ms INTEGER NOT NULL DEFAULT 0,
      lines_seen INTEGER NOT NULL DEFAULT 0,
      words_seen INTEGER NOT NULL DEFAULT 0,
      tokens_seen INTEGER NOT NULL DEFAULT 0,
      cards_mined INTEGER NOT NULL DEFAULT 0,
      lookup_count INTEGER NOT NULL DEFAULT 0,
      lookup_hits INTEGER NOT NULL DEFAULT 0,
      pause_count INTEGER NOT NULL DEFAULT 0,
      pause_ms INTEGER NOT NULL DEFAULT 0,
      seek_forward_count INTEGER NOT NULL DEFAULT 0,
      seek_backward_count INTEGER NOT NULL DEFAULT 0,
      media_buffer_events INTEGER NOT NULL DEFAULT 0,
      CREATED_DATE INTEGER,
      LAST_UPDATE_DATE INTEGER,
      FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_session_events(
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
      LAST_UPDATE_DATE INTEGER,
      FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_daily_rollups(
      rollup_day INTEGER NOT NULL,
      video_id INTEGER,
      total_sessions INTEGER NOT NULL DEFAULT 0,
      total_active_min REAL NOT NULL DEFAULT 0,
      total_lines_seen INTEGER NOT NULL DEFAULT 0,
      total_words_seen INTEGER NOT NULL DEFAULT 0,
      total_tokens_seen INTEGER NOT NULL DEFAULT 0,
      total_cards INTEGER NOT NULL DEFAULT 0,
      cards_per_hour REAL,
      words_per_min REAL,
      lookup_hit_rate REAL,
      CREATED_DATE INTEGER,
      LAST_UPDATE_DATE INTEGER,
      PRIMARY KEY (rollup_day, video_id)
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_monthly_rollups(
      rollup_month INTEGER NOT NULL,
      video_id INTEGER,
      total_sessions INTEGER NOT NULL DEFAULT 0,
      total_active_min REAL NOT NULL DEFAULT 0,
      total_lines_seen INTEGER NOT NULL DEFAULT 0,
      total_words_seen INTEGER NOT NULL DEFAULT 0,
      total_tokens_seen INTEGER NOT NULL DEFAULT 0,
      total_cards INTEGER NOT NULL DEFAULT 0,
      CREATED_DATE INTEGER,
      LAST_UPDATE_DATE INTEGER,
      PRIMARY KEY (rollup_month, video_id)
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_words(
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
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_kanji(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      kanji TEXT,
      first_seen REAL,
      last_seen REAL,
      frequency INTEGER,
      UNIQUE(kanji)
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_subtitle_lines(
      line_id INTEGER PRIMARY KEY AUTOINCREMENT,
      session_id INTEGER NOT NULL,
      event_id INTEGER,
      video_id INTEGER NOT NULL,
      anime_id INTEGER,
      line_index INTEGER NOT NULL,
      segment_start_ms INTEGER,
      segment_end_ms INTEGER,
      text TEXT NOT NULL,
      CREATED_DATE INTEGER,
      LAST_UPDATE_DATE INTEGER,
      FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE,
      FOREIGN KEY(event_id) REFERENCES imm_session_events(event_id) ON DELETE SET NULL,
      FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE,
      FOREIGN KEY(anime_id) REFERENCES imm_anime(anime_id) ON DELETE SET NULL
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_word_line_occurrences(
      line_id INTEGER NOT NULL,
      word_id INTEGER NOT NULL,
      occurrence_count INTEGER NOT NULL,
      PRIMARY KEY(line_id, word_id),
      FOREIGN KEY(line_id) REFERENCES imm_subtitle_lines(line_id) ON DELETE CASCADE,
      FOREIGN KEY(word_id) REFERENCES imm_words(id) ON DELETE CASCADE
    );
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_kanji_line_occurrences(
      line_id INTEGER NOT NULL,
      kanji_id INTEGER NOT NULL,
      occurrence_count INTEGER NOT NULL,
      PRIMARY KEY(line_id, kanji_id),
      FOREIGN KEY(line_id) REFERENCES imm_subtitle_lines(line_id) ON DELETE CASCADE,
      FOREIGN KEY(kanji_id) REFERENCES imm_kanji(id) ON DELETE CASCADE
    );
  `);
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

  if (currentVersion?.schema_version === 1) {
    addColumnIfMissing(db, 'imm_videos', 'CREATED_DATE');
    addColumnIfMissing(db, 'imm_videos', 'LAST_UPDATE_DATE');
    addColumnIfMissing(db, 'imm_sessions', 'CREATED_DATE');
    addColumnIfMissing(db, 'imm_sessions', 'LAST_UPDATE_DATE');
    addColumnIfMissing(db, 'imm_session_telemetry', 'CREATED_DATE');
    addColumnIfMissing(db, 'imm_session_telemetry', 'LAST_UPDATE_DATE');
    addColumnIfMissing(db, 'imm_session_events', 'CREATED_DATE');
    addColumnIfMissing(db, 'imm_session_events', 'LAST_UPDATE_DATE');
    addColumnIfMissing(db, 'imm_daily_rollups', 'CREATED_DATE');
    addColumnIfMissing(db, 'imm_daily_rollups', 'LAST_UPDATE_DATE');
    addColumnIfMissing(db, 'imm_monthly_rollups', 'CREATED_DATE');
    addColumnIfMissing(db, 'imm_monthly_rollups', 'LAST_UPDATE_DATE');

    const nowMs = Date.now();
    db.prepare(
      `
        UPDATE imm_videos
        SET
          CREATED_DATE = COALESCE(CREATED_DATE, created_at_ms),
          LAST_UPDATE_DATE = COALESCE(LAST_UPDATE_DATE, created_at_ms)
      `,
    ).run();
    db.prepare(
      `
        UPDATE imm_sessions
        SET
          CREATED_DATE = COALESCE(CREATED_DATE, started_at_ms),
          LAST_UPDATE_DATE = COALESCE(LAST_UPDATE_DATE, created_at_ms)
      `,
    ).run();
    db.prepare(
      `
        UPDATE imm_session_telemetry
        SET
          CREATED_DATE = COALESCE(CREATED_DATE, sample_ms),
          LAST_UPDATE_DATE = COALESCE(LAST_UPDATE_DATE, sample_ms)
      `,
    ).run();
    db.prepare(
      `
        UPDATE imm_session_events
        SET
          CREATED_DATE = COALESCE(CREATED_DATE, ts_ms),
          LAST_UPDATE_DATE = COALESCE(LAST_UPDATE_DATE, ts_ms)
      `,
    ).run();
    db.prepare(
      `
        UPDATE imm_daily_rollups
        SET
          CREATED_DATE = COALESCE(CREATED_DATE, ?),
          LAST_UPDATE_DATE = COALESCE(LAST_UPDATE_DATE, ?)
      `,
    ).run(nowMs, nowMs);
    db.prepare(
      `
        UPDATE imm_monthly_rollups
        SET
          CREATED_DATE = COALESCE(CREATED_DATE, ?),
          LAST_UPDATE_DATE = COALESCE(LAST_UPDATE_DATE, ?)
      `,
    ).run(nowMs, nowMs);
  }

  if (currentVersion?.schema_version === 1 || currentVersion?.schema_version === 2) {
    dropColumnIfExists(db, 'imm_videos', 'created_at_ms');
    dropColumnIfExists(db, 'imm_videos', 'updated_at_ms');
    dropColumnIfExists(db, 'imm_sessions', 'created_at_ms');
    dropColumnIfExists(db, 'imm_sessions', 'updated_at_ms');
  }

  if (currentVersion?.schema_version && currentVersion.schema_version < 5) {
    migrateLegacyAnimeMetadata(db);
  }

  if (currentVersion?.schema_version && currentVersion.schema_version < 6) {
    addColumnIfMissing(db, 'imm_words', 'part_of_speech', 'TEXT');
    addColumnIfMissing(db, 'imm_words', 'pos1', 'TEXT');
    addColumnIfMissing(db, 'imm_words', 'pos2', 'TEXT');
    addColumnIfMissing(db, 'imm_words', 'pos3', 'TEXT');
  }

  if (currentVersion?.schema_version && currentVersion.schema_version < 7) {
    db.exec(`
      CREATE TABLE IF NOT EXISTS imm_subtitle_lines(
        line_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        event_id INTEGER,
        video_id INTEGER NOT NULL,
        anime_id INTEGER,
        line_index INTEGER NOT NULL,
        segment_start_ms INTEGER,
        segment_end_ms INTEGER,
        text TEXT NOT NULL,
        CREATED_DATE INTEGER,
        LAST_UPDATE_DATE INTEGER,
        FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE,
        FOREIGN KEY(event_id) REFERENCES imm_session_events(event_id) ON DELETE SET NULL,
        FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE,
        FOREIGN KEY(anime_id) REFERENCES imm_anime(anime_id) ON DELETE SET NULL
      )
    `);
    db.exec(`
      CREATE TABLE IF NOT EXISTS imm_word_line_occurrences(
        line_id INTEGER NOT NULL,
        word_id INTEGER NOT NULL,
        occurrence_count INTEGER NOT NULL,
        PRIMARY KEY(line_id, word_id),
        FOREIGN KEY(line_id) REFERENCES imm_subtitle_lines(line_id) ON DELETE CASCADE,
        FOREIGN KEY(word_id) REFERENCES imm_words(id) ON DELETE CASCADE
      )
    `);
    db.exec(`
      CREATE TABLE IF NOT EXISTS imm_kanji_line_occurrences(
        line_id INTEGER NOT NULL,
        kanji_id INTEGER NOT NULL,
        occurrence_count INTEGER NOT NULL,
        PRIMARY KEY(line_id, kanji_id),
        FOREIGN KEY(line_id) REFERENCES imm_subtitle_lines(line_id) ON DELETE CASCADE,
        FOREIGN KEY(kanji_id) REFERENCES imm_kanji(id) ON DELETE CASCADE
      )
    `);
  }

  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_anime_normalized_title
    ON imm_anime(normalized_title_key)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_anime_anilist_id
    ON imm_anime(anilist_id)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_videos_anime_id
    ON imm_videos(anime_id)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_sessions_video_started
    ON imm_sessions(video_id, started_at_ms DESC)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_sessions_status_started
    ON imm_sessions(status, started_at_ms DESC)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_telemetry_session_sample
    ON imm_session_telemetry(session_id, sample_ms DESC)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_events_session_ts
    ON imm_session_events(session_id, ts_ms DESC)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_events_type_ts
    ON imm_session_events(event_type, ts_ms DESC)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_rollups_day_video
    ON imm_daily_rollups(rollup_day, video_id)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_rollups_month_video
    ON imm_monthly_rollups(rollup_month, video_id)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_words_headword_word_reading
    ON imm_words(headword, word, reading)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_kanji_kanji
    ON imm_kanji(kanji)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_subtitle_lines_session_line
    ON imm_subtitle_lines(session_id, line_index)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_subtitle_lines_video_line
    ON imm_subtitle_lines(video_id, line_index)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_subtitle_lines_anime_line
    ON imm_subtitle_lines(anime_id, line_index)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_word_line_occurrences_word
    ON imm_word_line_occurrences(word_id, line_id)
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_kanji_line_occurrences_kanji
    ON imm_kanji_line_occurrences(kanji_id, line_id)
  `);

  if (currentVersion?.schema_version && currentVersion.schema_version < SCHEMA_VERSION) {
    db.exec('DELETE FROM imm_daily_rollups');
    db.exec('DELETE FROM imm_monthly_rollups');
    db.exec(`UPDATE imm_rollup_state SET state_value = 0 WHERE state_key = 'last_rollup_sample_ms'`);
  }

  db.exec(`
    INSERT INTO imm_schema_version(schema_version, applied_at_ms)
    VALUES (${SCHEMA_VERSION}, ${Date.now()})
    ON CONFLICT DO NOTHING
  `);
}

export function createTrackerPreparedStatements(db: DatabaseSync): TrackerPreparedStatements {
  return {
    telemetryInsertStmt: db.prepare(`
      INSERT INTO imm_session_telemetry (
        session_id, sample_ms, total_watched_ms, active_watched_ms,
        lines_seen, words_seen, tokens_seen, cards_mined, lookup_count,
        lookup_hits, pause_count, pause_ms, seek_forward_count,
        seek_backward_count, media_buffer_events, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
      )
    `),
    eventInsertStmt: db.prepare(`
      INSERT INTO imm_session_events (
        session_id, ts_ms, event_type, line_index, segment_start_ms, segment_end_ms,
        words_delta, cards_delta, payload_json, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
      )
    `),
    wordUpsertStmt: db.prepare(`
      INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
      ) VALUES (
        ?, ?, ?, ?, ?, ?, ?, ?, ?, 1
      )
      ON CONFLICT(headword, word, reading) DO UPDATE SET
        frequency = COALESCE(frequency, 0) + 1,
        part_of_speech = CASE
          WHEN COALESCE(NULLIF(imm_words.part_of_speech, ''), 'other') = 'other'
            AND COALESCE(NULLIF(excluded.part_of_speech, ''), '') <> ''
          THEN excluded.part_of_speech
          ELSE imm_words.part_of_speech
        END,
        pos1 = COALESCE(NULLIF(imm_words.pos1, ''), excluded.pos1),
        pos2 = COALESCE(NULLIF(imm_words.pos2, ''), excluded.pos2),
        pos3 = COALESCE(NULLIF(imm_words.pos3, ''), excluded.pos3),
        first_seen = MIN(COALESCE(first_seen, excluded.first_seen), excluded.first_seen),
        last_seen = MAX(COALESCE(last_seen, excluded.last_seen), excluded.last_seen)
    `),
    kanjiUpsertStmt: db.prepare(`
      INSERT INTO imm_kanji (
        kanji, first_seen, last_seen, frequency
      ) VALUES (
        ?, ?, ?, 1
      )
      ON CONFLICT(kanji) DO UPDATE SET
        frequency = COALESCE(frequency, 0) + 1,
        first_seen = MIN(COALESCE(first_seen, excluded.first_seen), excluded.first_seen),
        last_seen = MAX(COALESCE(last_seen, excluded.last_seen), excluded.last_seen)
    `),
    subtitleLineInsertStmt: db.prepare(`
      INSERT INTO imm_subtitle_lines (
        session_id, event_id, video_id, anime_id, line_index, segment_start_ms,
        segment_end_ms, text, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
      )
    `),
    wordIdSelectStmt: db.prepare(`
      SELECT id FROM imm_words
      WHERE headword = ? AND word = ? AND reading = ?
    `),
    kanjiIdSelectStmt: db.prepare(`
      SELECT id FROM imm_kanji
      WHERE kanji = ?
    `),
    wordLineOccurrenceUpsertStmt: db.prepare(`
      INSERT INTO imm_word_line_occurrences (
        line_id, word_id, occurrence_count
      ) VALUES (
        ?, ?, ?
      )
      ON CONFLICT(line_id, word_id) DO UPDATE SET
        occurrence_count = imm_word_line_occurrences.occurrence_count + excluded.occurrence_count
    `),
    kanjiLineOccurrenceUpsertStmt: db.prepare(`
      INSERT INTO imm_kanji_line_occurrences (
        line_id, kanji_id, occurrence_count
      ) VALUES (
        ?, ?, ?
      )
      ON CONFLICT(line_id, kanji_id) DO UPDATE SET
        occurrence_count = imm_kanji_line_occurrences.occurrence_count + excluded.occurrence_count
    `),
    videoAnimeIdSelectStmt: db.prepare(`
      SELECT anime_id FROM imm_videos
      WHERE video_id = ?
    `),
  };
}

function incrementWordAggregate(
  stmts: TrackerPreparedStatements,
  occurrence: Extract<QueuedWrite, { kind: 'subtitleLine' }>['wordOccurrences'][number],
  firstSeen: number,
  lastSeen: number,
): number {
  for (let i = 0; i < occurrence.occurrenceCount; i += 1) {
    stmts.wordUpsertStmt.run(
      occurrence.headword,
      occurrence.word,
      occurrence.reading,
      occurrence.partOfSpeech,
      occurrence.pos1,
      occurrence.pos2,
      occurrence.pos3,
      firstSeen,
      lastSeen,
    );
  }
  const row = stmts.wordIdSelectStmt.get(
    occurrence.headword,
    occurrence.word,
    occurrence.reading,
  ) as { id: number } | null;
  if (!row?.id) {
    throw new Error(`Failed to resolve imm_words id for ${occurrence.headword}`);
  }
  return row.id;
}

function incrementKanjiAggregate(
  stmts: TrackerPreparedStatements,
  occurrence: Extract<QueuedWrite, { kind: 'subtitleLine' }>['kanjiOccurrences'][number],
  firstSeen: number,
  lastSeen: number,
): number {
  for (let i = 0; i < occurrence.occurrenceCount; i += 1) {
    stmts.kanjiUpsertStmt.run(occurrence.kanji, firstSeen, lastSeen);
  }
  const row = stmts.kanjiIdSelectStmt.get(occurrence.kanji) as { id: number } | null;
  if (!row?.id) {
    throw new Error(`Failed to resolve imm_kanji id for ${occurrence.kanji}`);
  }
  return row.id;
}

export function executeQueuedWrite(write: QueuedWrite, stmts: TrackerPreparedStatements): void {
  if (write.kind === 'telemetry') {
    stmts.telemetryInsertStmt.run(
      write.sessionId,
      write.sampleMs!,
      write.totalWatchedMs!,
      write.activeWatchedMs!,
      write.linesSeen!,
      write.wordsSeen!,
      write.tokensSeen!,
      write.cardsMined!,
      write.lookupCount!,
      write.lookupHits!,
      write.pauseCount!,
      write.pauseMs!,
      write.seekForwardCount!,
      write.seekBackwardCount!,
      write.mediaBufferEvents!,
      Date.now(),
      Date.now(),
    );
    return;
  }
  if (write.kind === 'word') {
    stmts.wordUpsertStmt.run(
      write.headword,
      write.word,
      write.reading,
      write.partOfSpeech,
      write.pos1,
      write.pos2,
      write.pos3,
      write.firstSeen,
      write.lastSeen,
    );
    return;
  }
  if (write.kind === 'kanji') {
    stmts.kanjiUpsertStmt.run(write.kanji, write.firstSeen, write.lastSeen);
    return;
  }
  if (write.kind === 'subtitleLine') {
    const animeRow = stmts.videoAnimeIdSelectStmt.get(write.videoId) as { anime_id: number | null } | null;
    const lineResult = stmts.subtitleLineInsertStmt.run(
      write.sessionId,
      null,
      write.videoId,
      animeRow?.anime_id ?? null,
      write.lineIndex,
      write.segmentStartMs ?? null,
      write.segmentEndMs ?? null,
      write.text,
      Date.now(),
      Date.now(),
    );
    const lineId = Number(lineResult.lastInsertRowid);
    for (const occurrence of write.wordOccurrences) {
      const wordId = incrementWordAggregate(stmts, occurrence, write.firstSeen, write.lastSeen);
      stmts.wordLineOccurrenceUpsertStmt.run(lineId, wordId, occurrence.occurrenceCount);
    }
    for (const occurrence of write.kanjiOccurrences) {
      const kanjiId = incrementKanjiAggregate(stmts, occurrence, write.firstSeen, write.lastSeen);
      stmts.kanjiLineOccurrenceUpsertStmt.run(lineId, kanjiId, occurrence.occurrenceCount);
    }
    return;
  }

  stmts.eventInsertStmt.run(
    write.sessionId,
    write.sampleMs!,
    write.eventType!,
    write.lineIndex ?? null,
    write.segmentStartMs ?? null,
    write.segmentEndMs ?? null,
    write.wordsDelta ?? 0,
    write.cardsDelta ?? 0,
    write.payloadJson ?? null,
    Date.now(),
    Date.now(),
  );
}

export function getOrCreateVideoRecord(
  db: DatabaseSync,
  videoKey: string,
  details: {
    canonicalTitle: string;
    sourcePath: string | null;
    sourceUrl: string | null;
    sourceType: number;
  },
): number {
  const existing = db
    .prepare('SELECT video_id FROM imm_videos WHERE video_key = ?')
    .get(videoKey) as { video_id: number } | null;
  if (existing?.video_id) {
    db.prepare(
      `
        UPDATE imm_videos
        SET
          canonical_title = ?,
          LAST_UPDATE_DATE = ?
        WHERE video_id = ?
      `,
    ).run(details.canonicalTitle || 'unknown', Date.now(), existing.video_id);
    return existing.video_id;
  }

  const nowMs = Date.now();
  const insert = db.prepare(`
    INSERT INTO imm_videos (
      video_key, canonical_title, source_type, source_path, source_url,
      duration_ms, file_size_bytes, codec_id, container_id, width_px, height_px,
      fps_x100, bitrate_kbps, audio_codec_id, hash_sha256, screenshot_path,
      metadata_json, CREATED_DATE, LAST_UPDATE_DATE
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const result = insert.run(
    videoKey,
    details.canonicalTitle || 'unknown',
    details.sourceType,
    details.sourcePath,
    details.sourceUrl,
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
  return Number(result.lastInsertRowid);
}

export function updateVideoMetadataRecord(
  db: DatabaseSync,
  videoId: number,
  metadata: VideoMetadata,
): void {
  db.prepare(
    `
      UPDATE imm_videos
      SET
        duration_ms = ?,
        file_size_bytes = ?,
        codec_id = ?,
        container_id = ?,
        width_px = ?,
        height_px = ?,
        fps_x100 = ?,
        bitrate_kbps = ?,
        audio_codec_id = ?,
        hash_sha256 = ?,
        screenshot_path = ?,
        metadata_json = ?,
        LAST_UPDATE_DATE = ?
      WHERE video_id = ?
    `,
  ).run(
    metadata.durationMs,
    metadata.fileSizeBytes,
    metadata.codecId,
    metadata.containerId,
    metadata.widthPx,
    metadata.heightPx,
    metadata.fpsX100,
    metadata.bitrateKbps,
    metadata.audioCodecId,
    metadata.hashSha256,
    metadata.screenshotPath,
    metadata.metadataJson,
    Date.now(),
    videoId,
  );
}

export function updateVideoTitleRecord(
  db: DatabaseSync,
  videoId: number,
  canonicalTitle: string,
): void {
  db.prepare(
    `
      UPDATE imm_videos
      SET
        canonical_title = ?,
        LAST_UPDATE_DATE = ?
      WHERE video_id = ?
    `,
  ).run(canonicalTitle, Date.now(), videoId);
}
