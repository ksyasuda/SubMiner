import { Database } from '../sqlite.js';
import type { DatabaseSync } from '../sqlite.js';
import {
  applyPragmas,
  ensureSchema,
  getOrCreateAnimeRecord,
  getOrCreateVideoRecord,
  linkVideoToAnimeRecord,
} from '../storage.js';
import { startSessionRecord } from '../session.js';
import { toDbTimestamp } from '../query-shared.js';

const SOURCE_TYPE_LOCAL = 1;
export const DAY_MS = 86_400_000;
// Noon UTC keeps every seeded timestamp on the same local day regardless of
// the timezone the test host runs in.
export const BASE_MS = Date.UTC(2026, 0, 5, 12, 0, 0);

export function createDb(): DatabaseSync {
  const db = new Database(':memory:');
  applyPragmas(db);
  ensureSchema(db);
  return db;
}

export function seedAnime(db: DatabaseSync, title: string, episodesTotal: number | null): number {
  const animeId = getOrCreateAnimeRecord(db, {
    parsedTitle: title,
    canonicalTitle: title,
    anilistId: null,
    titleRomaji: null,
    titleEnglish: null,
    titleNative: null,
    metadataJson: null,
  });
  if (episodesTotal !== null) {
    db.prepare('UPDATE imm_anime SET episodes_total = ? WHERE anime_id = ?').run(
      episodesTotal,
      animeId,
    );
  }
  return animeId;
}

export function seedVideo(
  db: DatabaseSync,
  animeId: number | null,
  name: string,
  options: { watched?: boolean } = {},
): number {
  const videoId = getOrCreateVideoRecord(db, `local:/tmp/${name}.mkv`, {
    canonicalTitle: name,
    sourcePath: `/tmp/${name}.mkv`,
    sourceUrl: null,
    sourceType: SOURCE_TYPE_LOCAL,
  });
  if (animeId !== null) {
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: `${name}.mkv`,
      parsedTitle: name,
      parsedSeason: 1,
      parsedEpisode: 1,
      parserSource: 'test',
      parserConfidence: 1,
      parseMetadataJson: null,
    });
  }
  if (options.watched) {
    db.prepare('UPDATE imm_videos SET watched = 1 WHERE video_id = ?').run(videoId);
  }
  return videoId;
}

export function seedEndedSession(
  db: DatabaseSync,
  videoId: number,
  startedAtMs: number,
  metrics: { activeMs: number; cards?: number; lines?: number; tokens?: number },
): number {
  const sessionId = startSessionRecord(db, videoId, startedAtMs).sessionId;
  db.prepare(
    `
    UPDATE imm_sessions SET
      ended_at_ms = ?,
      active_watched_ms = ?,
      total_watched_ms = ?,
      cards_mined = ?,
      lines_seen = ?,
      tokens_seen = ?
    WHERE session_id = ?
    `,
  ).run(
    toDbTimestamp(startedAtMs + metrics.activeMs),
    metrics.activeMs,
    metrics.activeMs,
    metrics.cards ?? 0,
    metrics.lines ?? 0,
    metrics.tokens ?? 0,
    sessionId,
  );
  return sessionId;
}

// libsql attaches a per-query `_metadata` property to result rows; strip it so
// row snapshots can be compared with deepEqual.
export function cleanRow<T>(row: unknown): T {
  const { _metadata: _ignored, ...rest } = row as Record<string, unknown>;
  return rest as T;
}

export interface GlobalSnapshot {
  total_sessions: number;
  total_active_ms: number;
  total_cards: number;
  active_days: number;
  episodes_started: number;
  episodes_completed: number;
  anime_completed: number;
}

export function snapshotGlobal(db: DatabaseSync): GlobalSnapshot {
  const row = db
    .prepare(
      `SELECT total_sessions, total_active_ms, total_cards, active_days,
              episodes_started, episodes_completed, anime_completed
       FROM imm_lifetime_global WHERE global_id = 1`,
    )
    .get();
  return cleanRow<GlobalSnapshot>(row);
}

export function snapshotMedia(db: DatabaseSync): unknown[] {
  return db
    .prepare(
      `SELECT video_id, total_sessions, total_active_ms, total_cards,
              total_lines_seen, total_tokens_seen, completed,
              CAST(first_watched_ms AS REAL) AS first_watched,
              CAST(last_watched_ms AS REAL) AS last_watched
       FROM imm_lifetime_media ORDER BY video_id`,
    )
    .all()
    .map((row) => cleanRow(row));
}

export function snapshotAnime(db: DatabaseSync): unknown[] {
  return db
    .prepare(
      `SELECT anime_id, total_sessions, total_active_ms, total_cards,
              total_lines_seen, total_tokens_seen, episodes_started, episodes_completed,
              CAST(first_watched_ms AS REAL) AS first_watched,
              CAST(last_watched_ms AS REAL) AS last_watched
       FROM imm_lifetime_anime ORDER BY anime_id`,
    )
    .all()
    .map((row) => cleanRow(row));
}
