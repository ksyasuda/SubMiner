import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Database } from 'bun:sqlite';
import { resolveConfigDir } from '../src/config/path-resolution.js';
import { readLauncherMainConfigObject } from './config/shared-config-reader.js';
import type { HistoryVideoRow } from './history-types.js';
import { resolvePathMaybe } from './util.js';

export function resolveImmersionDbPath(): string {
  const root = readLauncherMainConfigObject();
  const tracking =
    root?.immersionTracking &&
    typeof root.immersionTracking === 'object' &&
    !Array.isArray(root.immersionTracking)
      ? (root.immersionTracking as Record<string, unknown>)
      : null;
  const configured = typeof tracking?.dbPath === 'string' ? tracking.dbPath.trim() : '';
  if (configured) return resolvePathMaybe(configured);

  const configDir = resolveConfigDir({
    platform: process.platform,
    appDataDir: process.env.APPDATA,
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    existsSync: fs.existsSync,
  });
  return path.join(configDir, 'immersion.sqlite');
}

interface RawHistoryRow {
  video_id: number;
  source_path: string | null;
  parsed_title: string | null;
  parsed_season: number | null;
  parsed_episode: number | null;
  anime_title: string | null;
  last_watched_ms: number | bigint | null;
  cover_blob_hash: string | null;
}

export function queryLocalWatchHistory(dbPath: string): HistoryVideoRow[] {
  try {
    return readHistoryRows(dbPath, { readonly: true });
  } catch (error) {
    if (!isReadonlyWalRetryError(error, dbPath)) throw error;
    return readHistoryRows(dbPath, { readwrite: true, create: false });
  }
}

export function isReadonlyWalRetryError(error: unknown, dbPath: string): boolean {
  if (!isWalModeSqliteDatabase(dbPath)) return false;
  const code =
    typeof error === 'object' && error !== null && 'code' in error
      ? String((error as { code?: unknown }).code ?? '')
      : '';
  const message = error instanceof Error ? error.message : String(error);
  const text = `${code} ${message}`.toLowerCase();
  return (
    text.includes('readonly') ||
    text.includes('read-only') ||
    text.includes('attempt to write a readonly database') ||
    text.includes('sqlite_cantopen') ||
    text.includes('unable to open database file')
  );
}

function isWalModeSqliteDatabase(dbPath: string): boolean {
  const header = Buffer.alloc(20);
  let fd: number | null = null;
  try {
    fd = fs.openSync(dbPath, 'r');
    if (fs.readSync(fd, header, 0, header.length, 0) < header.length) return false;
  } catch {
    return false;
  } finally {
    if (fd !== null) fs.closeSync(fd);
  }
  return header.subarray(0, 16).toString('ascii') === 'SQLite format 3\0' && header[18] === 2;
}

function tableExists(db: Database, tableName: string): boolean {
  return Boolean(
    db.query(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(tableName),
  );
}

function readHistoryRows(
  dbPath: string,
  options: { readonly?: boolean; readwrite?: boolean; create?: boolean },
): HistoryVideoRow[] {
  const db = new Database(dbPath, options);
  try {
    const hasMediaArt = tableExists(db, 'imm_media_art');
    const coverSelect = hasMediaArt
      ? `COALESCE(
          ma.cover_blob_hash,
          (SELECT ma2.cover_blob_hash
           FROM imm_media_art ma2
           JOIN imm_videos v2 ON v2.video_id = ma2.video_id
           WHERE v2.anime_id = v.anime_id AND ma2.cover_blob_hash IS NOT NULL
           LIMIT 1)
        ) AS cover_blob_hash`
      : 'NULL AS cover_blob_hash';
    const coverJoin = hasMediaArt ? 'LEFT JOIN imm_media_art ma ON ma.video_id = v.video_id' : '';
    const rows = db
      .query<RawHistoryRow>(
        `
        SELECT
          v.video_id,
          v.source_path,
          v.parsed_title,
          v.parsed_season,
          v.parsed_episode,
          COALESCE(a.title_romaji, a.canonical_title) AS anime_title,
          MAX(CAST(s.started_at_ms AS INTEGER)) AS last_watched_ms,
          ${coverSelect}
        FROM imm_sessions s
        JOIN imm_videos v ON v.video_id = s.video_id
        LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
        ${coverJoin}
        WHERE v.source_type = 1 AND v.source_path IS NOT NULL AND v.source_path != ''
        GROUP BY v.video_id
        ORDER BY last_watched_ms DESC
        `,
      )
      .all();

    return rows
      .filter((row) => typeof row.source_path === 'string' && row.source_path.length > 0)
      .map((row) => ({
        videoId: row.video_id,
        sourcePath: row.source_path!,
        parsedTitle: row.parsed_title,
        parsedSeason: row.parsed_season,
        parsedEpisode: row.parsed_episode,
        animeTitle: row.anime_title,
        lastWatchedMs: Number(row.last_watched_ms ?? 0),
        coverBlobHash: row.cover_blob_hash,
      }));
  } finally {
    db.close();
  }
}
