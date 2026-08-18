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

/**
 * Snapshots for long series run to hundreds of MB each, so everything here reads them off the main
 * thread's critical path: file IO is async and only the unavoidable JSON.parse runs on the loop,
 * one file at a time. Reading the whole directory synchronously used to block the process for
 * multiple seconds, long enough for the compositor to declare the app unresponsive mid-playback.
 */
export async function readCachedSnapshots(
  outputDir: string,
): Promise<CharacterDictionarySnapshot[]> {
  let entries: fs.Dirent[] = [];
  try {
    entries = await fs.promises.readdir(getSnapshotsDir(outputDir), { withFileTypes: true });
  } catch {
    return [];
  }

  const names = entries
    .filter((entry) => entry.isFile() && /^anilist-\d+\.json$/.test(entry.name))
    .map((entry) => entry.name)
    .sort((left, right) => left.localeCompare(right));

  const snapshots: CharacterDictionarySnapshot[] = [];
  for (const name of names) {
    const snapshot = await readSnapshot(path.join(getSnapshotsDir(outputDir), name));
    if (snapshot) {
      snapshots.push(snapshot);
    }
  }
  return snapshots;
}

export async function readSnapshot(
  snapshotPath: string,
): Promise<CharacterDictionarySnapshot | null> {
  try {
    const raw = await fs.promises.readFile(snapshotPath, 'utf8');
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

// Flushing in a few-MB batches keeps each stringify-and-write slice short; a single
// JSON.stringify of a large snapshot blocks the event loop for seconds.
const SNAPSHOT_WRITE_FLUSH_BYTES = 4 * 1024 * 1024;

// Distinguishes concurrent writes of the same snapshot within one process; the pid alone only
// separates processes, so two overlapping writers would otherwise stream into the same temp file.
let snapshotWriteSequence = 0;

/**
 * Streams the snapshot to disk piece by piece instead of stringifying it in one shot, then renames
 * the finished file into place so a crash mid-write (or two concurrent writers for the same media)
 * can never leave a torn file where a snapshot used to be.
 */
export async function writeSnapshot(
  snapshotPath: string,
  snapshot: CharacterDictionarySnapshot,
): Promise<void> {
  ensureDir(path.dirname(snapshotPath));
  snapshotWriteSequence += 1;
  const tempPath = `${snapshotPath}.tmp-${process.pid}-${snapshotWriteSequence}`;
  const handle = await fs.promises.open(tempPath, 'w');
  try {
    let buffered: string[] = [];
    let bufferedBytes = 0;
    const push = async (chunk: string): Promise<void> => {
      buffered.push(chunk);
      bufferedBytes += chunk.length;
      if (bufferedBytes >= SNAPSHOT_WRITE_FLUSH_BYTES) {
        const joined = buffered.join('');
        buffered = [];
        bufferedBytes = 0;
        await handle.write(joined, null, 'utf8');
      }
    };
    const writeArray = async (key: string, items: readonly unknown[]): Promise<void> => {
      await push(`,${JSON.stringify(key)}:[`);
      for (let i = 0; i < items.length; i += 1) {
        await push(`${i > 0 ? ',' : ''}${JSON.stringify(items[i])}`);
      }
      await push(']');
    };

    const { termEntries, images, ...scalars } = snapshot;
    const head = JSON.stringify(scalars);
    await push(head.slice(0, -1));
    await writeArray('termEntries', termEntries);
    await writeArray('images', images);
    await push('}');
    if (buffered.length > 0) {
      await handle.write(buffered.join(''), null, 'utf8');
    }
  } catch (error) {
    await handle.close();
    await fs.promises.rm(tempPath, { force: true });
    throw error;
  }
  await handle.close();
  await fs.promises.rename(tempPath, snapshotPath);
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
