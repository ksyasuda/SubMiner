import type { DatabaseSync } from 'node:sqlite';
import { SCHEMA_VERSION } from './types';
import type { QueuedWrite, VideoMetadata } from './types';

export interface TrackerPreparedStatements {
  telemetryInsertStmt: ReturnType<DatabaseSync['prepare']>;
  eventInsertStmt: ReturnType<DatabaseSync['prepare']>;
}

export function applyPragmas(db: DatabaseSync): void {
  db.exec('PRAGMA journal_mode = WAL');
  db.exec('PRAGMA synchronous = NORMAL');
  db.exec('PRAGMA foreign_keys = ON');
  db.exec('PRAGMA busy_timeout = 2500');
}

export function ensureSchema(db: DatabaseSync): void {
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_schema_version (
      schema_version INTEGER PRIMARY KEY,
      applied_at_ms INTEGER NOT NULL
    );
  `);

  const currentVersion = db
    .prepare('SELECT schema_version FROM imm_schema_version ORDER BY schema_version DESC LIMIT 1')
    .get() as { schema_version: number } | null;
  if (currentVersion?.schema_version === SCHEMA_VERSION) {
    return;
  }

  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_videos(
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
      created_at_ms INTEGER NOT NULL, updated_at_ms INTEGER NOT NULL
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
      created_at_ms INTEGER NOT NULL, updated_at_ms INTEGER NOT NULL,
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
      PRIMARY KEY (rollup_month, video_id)
    );
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
        seek_backward_count, media_buffer_events
      ) VALUES (
        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
      )
    `),
    eventInsertStmt: db.prepare(`
      INSERT INTO imm_session_events (
        session_id, ts_ms, event_type, line_index, segment_start_ms, segment_end_ms,
        words_delta, cards_delta, payload_json
      ) VALUES (
        ?, ?, ?, ?, ?, ?, ?, ?, ?
      )
    `),
  };
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
    );
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
      'UPDATE imm_videos SET canonical_title = ?, updated_at_ms = ? WHERE video_id = ?',
    ).run(details.canonicalTitle || 'unknown', Date.now(), existing.video_id);
    return existing.video_id;
  }

  const nowMs = Date.now();
  const insert = db.prepare(`
    INSERT INTO imm_videos (
      video_key, canonical_title, source_type, source_path, source_url,
      duration_ms, file_size_bytes, codec_id, container_id, width_px, height_px,
      fps_x100, bitrate_kbps, audio_codec_id, hash_sha256, screenshot_path,
      metadata_json, created_at_ms, updated_at_ms
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
        updated_at_ms = ?
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
  db.prepare('UPDATE imm_videos SET canonical_title = ?, updated_at_ms = ? WHERE video_id = ?').run(
    canonicalTitle,
    Date.now(),
    videoId,
  );
}
