import fs from 'node:fs';
import path from 'node:path';
import { parseMediaInfo } from '../src/jimaku/utils.js';
import { collectVideos } from './picker.js';
import type { HistorySeriesEntry, HistoryVideoRow, SeasonDirEntry } from './history-types.js';

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

export function groupHistoryBySeries(
  rows: HistoryVideoRow[],
  existsFn: (candidate: string) => boolean = fs.existsSync,
): HistorySeriesEntry[] {
  const byRoot = new Map<string, HistorySeriesEntry>();
  const sorted = [...rows].sort((a, b) => b.lastWatchedMs - a.lastWatchedMs);

  for (const row of sorted) {
    const seriesRoot = resolveSeriesRoot(row.sourcePath);
    const existing = byRoot.get(seriesRoot);
    if (existing) {
      if (existing.coverBlobHash === null && row.coverBlobHash !== null) {
        existing.coverBlobHash = row.coverBlobHash;
      }
      continue;
    }
    if (!existsFn(seriesRoot)) continue;
    const displayName =
      row.parsedTitle?.trim() || row.animeTitle?.trim() || path.basename(seriesRoot);
    byRoot.set(seriesRoot, {
      seriesRoot,
      displayName,
      lastWatched: row,
      coverBlobHash: row.coverBlobHash,
    });
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

function findFirstEpisodeInNextSeason(resolvedLast: string, dir: string): string | null {
  const seriesRoot = resolveSeriesRoot(resolvedLast);
  if (seriesRoot === dir) return null;
  const seasons = listSeasonDirs(seriesRoot);
  const currentIdx = seasons.findIndex((season) => path.resolve(season.path) === dir);
  if (currentIdx < 0 || currentIdx + 1 >= seasons.length) return null;
  const nextSeason = sortVideosByEpisode(collectVideos(seasons[currentIdx + 1]!.path, false));
  return nextSeason[0] ?? null;
}

export function findNextEpisode(lastPath: string): string | null {
  const resolvedLast = path.resolve(lastPath);
  const dir = path.dirname(resolvedLast);
  const episodes = sortVideosByEpisode(collectVideos(dir, false));
  const idx = episodes.indexOf(resolvedLast);

  if (idx >= 0) {
    if (idx + 1 < episodes.length) return episodes[idx + 1]!;
  } else {
    const lastInfo = parseMediaInfo(resolvedLast);
    if (lastInfo.episode !== null) {
      const candidate = episodes.find((episode) => {
        const info = parseMediaInfo(episode);
        return info.episode !== null && info.episode > lastInfo.episode!;
      });
      if (candidate) return candidate;
    }
  }

  return findFirstEpisodeInNextSeason(resolvedLast, dir);
}

function findLastEpisodeInPreviousSeason(resolvedCurrent: string, dir: string): string | null {
  const seriesRoot = resolveSeriesRoot(resolvedCurrent);
  if (seriesRoot === dir) return null;
  const seasons = listSeasonDirs(seriesRoot);
  const currentIdx = seasons.findIndex((season) => path.resolve(season.path) === dir);
  if (currentIdx <= 0) return null;
  const previousSeason = sortVideosByEpisode(collectVideos(seasons[currentIdx - 1]!.path, false));
  return previousSeason[previousSeason.length - 1] ?? null;
}

export function findPreviousEpisode(currentPath: string): string | null {
  const resolvedCurrent = path.resolve(currentPath);
  const dir = path.dirname(resolvedCurrent);
  const episodes = sortVideosByEpisode(collectVideos(dir, false));
  const idx = episodes.indexOf(resolvedCurrent);

  if (idx >= 0) {
    if (idx - 1 >= 0) return episodes[idx - 1]!;
  } else {
    const currentInfo = parseMediaInfo(resolvedCurrent);
    if (currentInfo.episode !== null) {
      const candidates = episodes.filter((episode) => {
        const info = parseMediaInfo(episode);
        return info.episode !== null && info.episode < currentInfo.episode!;
      });
      const candidate = candidates[candidates.length - 1];
      if (candidate) return candidate;
    }
  }

  return findLastEpisodeInPreviousSeason(resolvedCurrent, dir);
}
