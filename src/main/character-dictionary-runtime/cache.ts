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
