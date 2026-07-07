import * as fs from 'fs';
import * as path from 'path';
import { createHash } from 'node:crypto';
import { CHARACTER_DICTIONARY_FORMAT_VERSION } from './constants';
import { ensureDir } from './fs-utils';
import type {
  CharacterDictionarySnapshot,
  CharacterDictionarySnapshotImage,
  CharacterDictionaryTermEntry,
} from './types';

function getSnapshotsDir(outputDir: string): string {
  return path.join(outputDir, 'snapshots');
}

export function getSnapshotPath(outputDir: string, mediaId: number): string {
  return path.join(getSnapshotsDir(outputDir), `anilist-${mediaId}.json`);
}

export function getMergedZipPath(outputDir: string): string {
  return path.join(outputDir, 'merged.zip');
}

type MediaResolutionCacheEntry = {
  seriesKey: string;
  mediaId: number;
  mediaTitle: string;
};

type MediaResolutionCacheFile = {
  entries?: MediaResolutionCacheEntry[];
};

function getMediaResolutionCachePath(outputDir: string): string {
  return path.join(outputDir, 'anilist-resolution-cache.json');
}

function normalizeMediaResolutionEntry(value: unknown): MediaResolutionCacheEntry | null {
  if (!value || typeof value !== 'object') return null;
  const raw = value as Partial<MediaResolutionCacheEntry>;
  const seriesKey = typeof raw.seriesKey === 'string' ? raw.seriesKey.trim() : '';
  const mediaTitle = typeof raw.mediaTitle === 'string' ? raw.mediaTitle.trim() : '';
  if (typeof raw.mediaId !== 'number' || !Number.isFinite(raw.mediaId)) return null;
  const mediaId = Math.floor(raw.mediaId);
  if (!seriesKey || mediaId <= 0 || !mediaTitle) return null;
  return {
    seriesKey,
    mediaId,
    mediaTitle,
  };
}

function readMediaResolutionEntries(outputDir: string): MediaResolutionCacheEntry[] {
  try {
    const parsed = JSON.parse(
      fs.readFileSync(getMediaResolutionCachePath(outputDir), 'utf8'),
    ) as MediaResolutionCacheFile;
    if (!Array.isArray(parsed.entries)) return [];
    const byKey = new Map<string, MediaResolutionCacheEntry>();
    for (const value of parsed.entries) {
      const normalized = normalizeMediaResolutionEntry(value);
      if (normalized) byKey.set(normalized.seriesKey, normalized);
    }
    return [...byKey.values()];
  } catch {
    return [];
  }
}

function writeMediaResolutionEntries(
  outputDir: string,
  entries: MediaResolutionCacheEntry[],
): void {
  ensureDir(outputDir);
  fs.writeFileSync(
    getMediaResolutionCachePath(outputDir),
    JSON.stringify({ entries }, null, 2),
    'utf8',
  );
}

export function readCachedMediaResolution(
  outputDir: string,
  seriesKey: string,
): MediaResolutionCacheEntry | null {
  const normalizedKey = seriesKey.trim();
  if (!normalizedKey) return null;
  return (
    readMediaResolutionEntries(outputDir).find((entry) => entry.seriesKey === normalizedKey) ?? null
  );
}

export function writeCachedMediaResolution(
  outputDir: string,
  entry: MediaResolutionCacheEntry,
): void {
  const normalized = normalizeMediaResolutionEntry(entry);
  if (!normalized) return;
  const remaining = readMediaResolutionEntries(outputDir).filter(
    (existing) => existing.seriesKey !== normalized.seriesKey,
  );
  writeMediaResolutionEntries(outputDir, [...remaining, normalized]);
}

export function readCachedSnapshots(outputDir: string): CharacterDictionarySnapshot[] {
  let entries: fs.Dirent[] = [];
  try {
    entries = fs.readdirSync(getSnapshotsDir(outputDir), { withFileTypes: true });
  } catch {
    return [];
  }

  return entries
    .filter((entry) => entry.isFile() && /^anilist-\d+\.json$/.test(entry.name))
    .sort((left, right) => left.name.localeCompare(right.name))
    .map((entry) => readSnapshot(path.join(getSnapshotsDir(outputDir), entry.name)))
    .filter((snapshot): snapshot is CharacterDictionarySnapshot => snapshot !== null);
}

export function readSnapshot(snapshotPath: string): CharacterDictionarySnapshot | null {
  try {
    const raw = fs.readFileSync(snapshotPath, 'utf8');
    const parsed = JSON.parse(raw) as Partial<CharacterDictionarySnapshot>;
    if (!parsed || typeof parsed !== 'object') {
      return null;
    }
    if (
      parsed.formatVersion !== CHARACTER_DICTIONARY_FORMAT_VERSION ||
      typeof parsed.mediaId !== 'number' ||
      typeof parsed.mediaTitle !== 'string' ||
      typeof parsed.entryCount !== 'number' ||
      typeof parsed.updatedAt !== 'number' ||
      !Array.isArray(parsed.termEntries) ||
      !Array.isArray(parsed.images)
    ) {
      return null;
    }
    return {
      formatVersion: parsed.formatVersion,
      mediaId: parsed.mediaId,
      mediaTitle: parsed.mediaTitle,
      entryCount: parsed.entryCount,
      updatedAt: parsed.updatedAt,
      nameSplitSource: parsed.nameSplitSource === 'mecab' ? 'mecab' : 'heuristic',
      termEntries: parsed.termEntries as CharacterDictionaryTermEntry[],
      images: parsed.images as CharacterDictionarySnapshotImage[],
    };
  } catch {
    return null;
  }
}

export function writeSnapshot(snapshotPath: string, snapshot: CharacterDictionarySnapshot): void {
  ensureDir(path.dirname(snapshotPath));
  fs.writeFileSync(snapshotPath, JSON.stringify(snapshot, null, 2), 'utf8');
}

export function buildMergedRevision(
  mediaIds: number[],
  snapshots: CharacterDictionarySnapshot[],
): string {
  const hash = createHash('sha1');
  hash.update(
    JSON.stringify({
      mediaIds,
      snapshots: snapshots.map((snapshot) => ({
        mediaId: snapshot.mediaId,
        updatedAt: snapshot.updatedAt,
        entryCount: snapshot.entryCount,
      })),
    }),
  );
  return hash.digest('hex').slice(0, 12);
}

export function normalizeMergedMediaIds(mediaIds: number[]): number[] {
  return [
    ...new Set(
      mediaIds
        .filter((mediaId) => Number.isFinite(mediaId) && mediaId > 0)
        .map((mediaId) => Math.floor(mediaId)),
    ),
  ].sort((left, right) => left - right);
}
