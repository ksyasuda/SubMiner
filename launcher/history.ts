import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Database } from 'bun:sqlite';
import { parseMediaInfo } from '../src/jimaku/utils.js';
import { resolveConfigDir } from '../src/config/path-resolution.js';
import { readLauncherMainConfigObject } from './config/shared-config-reader.js';
import { collectVideos } from './picker.js';
import { resolvePathMaybe } from './util.js';

export interface HistoryVideoRow {
  videoId: number;
  sourcePath: string;
  parsedTitle: string | null;
  parsedSeason: number | null;
  parsedEpisode: number | null;
  animeTitle: string | null;
  lastWatchedMs: number;
}

export interface HistorySeriesEntry {
  seriesRoot: string;
  displayName: string;
  lastWatched: HistoryVideoRow;
}

export interface SeasonDirEntry {
  name: string;
  path: string;
  season: number | null;
}

const SEASON_DIR_PATTERN = /^(?:season|s)[\s._-]*(\d{1,3})\b/i;

export function seasonNumberFromDirName(name: string): number | null {
  const match = name.trim().match(SEASON_DIR_PATTERN);
  if (!match) return null;
  const parsed = Number.parseInt(match[1]!, 10);
  return Number.isFinite(parsed) ? parsed : null;
}

export function resolveSeriesRoot(filePath: string): string {
  const parent = path.dirname(filePath);
  if (seasonNumberFromDirName(path.basename(parent)) !== null) {
    return path.dirname(parent);
  }
  return parent;
}

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
}

export function queryLocalWatchHistory(dbPath: string): HistoryVideoRow[] {
  try {
    return readHistoryRows(dbPath, { readonly: true });
  } catch {
    // The database uses WAL mode; a read-only connection cannot recreate the
    // -wal/-shm files after the app closed them, so retry with a read-write
    // handle (still only running SELECTs).
    return readHistoryRows(dbPath, { readwrite: true, create: false });
  }
}

function readHistoryRows(
  dbPath: string,
  options: { readonly?: boolean; readwrite?: boolean; create?: boolean },
): HistoryVideoRow[] {
  const db = new Database(dbPath, options);
  try {
    const rows = db
      .query(
        `
        SELECT
          v.video_id,
          v.source_path,
          v.parsed_title,
          v.parsed_season,
          v.parsed_episode,
          COALESCE(a.title_romaji, a.canonical_title) AS anime_title,
          MAX(CAST(s.started_at_ms AS INTEGER)) AS last_watched_ms
        FROM imm_sessions s
        JOIN imm_videos v ON v.video_id = s.video_id
        LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
        WHERE v.source_type = 1 AND v.source_path IS NOT NULL AND v.source_path != ''
        GROUP BY v.video_id
        ORDER BY last_watched_ms DESC
        `,
      )
      .all() as RawHistoryRow[];

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
      }));
  } finally {
    db.close();
  }
}

export function groupHistoryBySeries(
  rows: HistoryVideoRow[],
  existsFn: (candidate: string) => boolean = fs.existsSync,
): HistorySeriesEntry[] {
  const byRoot = new Map<string, HistorySeriesEntry>();
  const sorted = [...rows].sort((a, b) => b.lastWatchedMs - a.lastWatchedMs);

  for (const row of sorted) {
    const seriesRoot = resolveSeriesRoot(row.sourcePath);
    if (byRoot.has(seriesRoot)) continue;
    if (!existsFn(seriesRoot)) continue;
    const displayName =
      row.parsedTitle?.trim() || row.animeTitle?.trim() || path.basename(seriesRoot);
    byRoot.set(seriesRoot, { seriesRoot, displayName, lastWatched: row });
  }

  return Array.from(byRoot.values());
}

function compareNatural(a: string, b: string): number {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
}

export function sortVideosByEpisode(videos: string[]): string[] {
  const parsed = videos.map((video) => ({ video, info: parseMediaInfo(video) }));
  parsed.sort((a, b) => {
    if (a.info.episode !== null && b.info.episode !== null) {
      const seasonA = a.info.season ?? 0;
      const seasonB = b.info.season ?? 0;
      if (seasonA !== seasonB) return seasonA - seasonB;
      if (a.info.episode !== b.info.episode) return a.info.episode - b.info.episode;
    }
    return compareNatural(a.video, b.video);
  });
  return parsed.map((entry) => entry.video);
}

function dirContainsVideo(dir: string): boolean {
  return collectVideos(dir, true).length > 0;
}

export function listSeasonDirs(seriesRoot: string): SeasonDirEntry[] {
  let entries: fs.Dirent[];
  try {
    entries = fs.readdirSync(seriesRoot, { withFileTypes: true });
  } catch {
    return [];
  }

  const dirs = entries
    .filter((entry) => entry.isDirectory())
    .map((entry) => ({
      name: entry.name,
      path: path.join(seriesRoot, entry.name),
      season: seasonNumberFromDirName(entry.name),
    }))
    .filter((entry) => dirContainsVideo(entry.path));

  dirs.sort((a, b) => {
    if (a.season !== null && b.season !== null && a.season !== b.season) {
      return a.season - b.season;
    }
    return compareNatural(a.name, b.name);
  });
  return dirs;
}

export function findNextEpisode(lastPath: string): string | null {
  const resolvedLast = path.resolve(lastPath);
  const dir = path.dirname(resolvedLast);
  const episodes = sortVideosByEpisode(collectVideos(dir, false));
  const idx = episodes.indexOf(resolvedLast);

  if (idx >= 0) {
    if (idx + 1 < episodes.length) return episodes[idx + 1]!;
  } else {
    // Last-watched file no longer exists: fall back to the first episode with
    // a higher parsed episode number in the same directory.
    const lastInfo = parseMediaInfo(resolvedLast);
    if (lastInfo.episode !== null) {
      const candidate = episodes.find((episode) => {
        const info = parseMediaInfo(episode);
        return info.episode !== null && info.episode > lastInfo.episode!;
      });
      if (candidate) return candidate;
    }
    return null;
  }

  // End of the season: continue with the first episode of the next season.
  const seriesRoot = resolveSeriesRoot(resolvedLast);
  if (seriesRoot === dir) return null;
  const seasons = listSeasonDirs(seriesRoot);
  const currentIdx = seasons.findIndex((season) => path.resolve(season.path) === dir);
  if (currentIdx < 0 || currentIdx + 1 >= seasons.length) return null;
  const nextSeason = sortVideosByEpisode(collectVideos(seasons[currentIdx + 1]!.path, false));
  return nextSeason[0] ?? null;
}
