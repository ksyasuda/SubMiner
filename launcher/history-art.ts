import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Database } from 'bun:sqlite';
import { isReadonlyWalRetryError } from './history-db.js';

const COVER_EXTENSIONS = ['.jpg', '.png', '.webp', '.gif'] as const;

export function getDefaultCoverCacheDir(): string {
  return path.join(os.homedir(), '.cache', 'subminer', 'covers');
}

export function detectImageExtension(blob: Buffer): string {
  if (blob.length >= 8 && blob.subarray(0, 8).equals(Buffer.from('89504e470d0a1a0a', 'hex'))) {
    return '.png';
  }
  if (blob.length >= 3 && blob[0] === 0xff && blob[1] === 0xd8 && blob[2] === 0xff) {
    return '.jpg';
  }
  if (
    blob.length >= 12 &&
    blob.subarray(0, 4).toString('ascii') === 'RIFF' &&
    blob.subarray(8, 12).toString('ascii') === 'WEBP'
  ) {
    return '.webp';
  }
  if (blob.length >= 4 && blob.subarray(0, 3).toString('ascii') === 'GIF') {
    return '.gif';
  }
  return '.jpg';
}

function findCachedCover(cacheDir: string, hash: string): string | null {
  for (const ext of COVER_EXTENSIONS) {
    const candidate = path.join(cacheDir, `${hash}${ext}`);
    try {
      if (fs.statSync(candidate).size > 0) return candidate;
    } catch {
      // not cached with this extension
    }
  }
  return null;
}

function queryCoverBlobs(
  dbPath: string,
  hashes: string[],
  options: { readonly?: boolean; readwrite?: boolean; create?: boolean },
): Map<string, Buffer> {
  const blobs = new Map<string, Buffer>();
  const db = new Database(dbPath, options);
  try {
    const hasBlobTable = db
      .query(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'imm_cover_art_blobs'`)
      .get();
    if (!hasBlobTable) return blobs;

    const stmt = db.query<{ cover_blob: Uint8Array | null }>(
      'SELECT cover_blob FROM imm_cover_art_blobs WHERE blob_hash = ?',
    );
    for (const hash of hashes) {
      const row = stmt.get(hash);
      if (row?.cover_blob && row.cover_blob.length > 0) {
        blobs.set(hash, Buffer.from(row.cover_blob));
      }
    }
    return blobs;
  } finally {
    db.close();
  }
}

/**
 * Ensures cover art blobs referenced by hash exist as image files in the cache
 * directory, extracting missing ones from the stats database. Returns a map of
 * blob hash to on-disk image path for every cover that could be materialized.
 */
export function materializeCoverArt(
  dbPath: string,
  hashes: Array<string | null | undefined>,
  cacheDir: string = getDefaultCoverCacheDir(),
): Map<string, string> {
  const wanted = Array.from(new Set(hashes.filter((hash): hash is string => Boolean(hash))));
  const resolved = new Map<string, string>();
  if (wanted.length === 0) return resolved;

  const missing: string[] = [];
  for (const hash of wanted) {
    const cached = findCachedCover(cacheDir, hash);
    if (cached) {
      resolved.set(hash, cached);
    } else {
      missing.push(hash);
    }
  }
  if (missing.length === 0) return resolved;

  let blobs: Map<string, Buffer>;
  try {
    try {
      blobs = queryCoverBlobs(dbPath, missing, { readonly: true });
    } catch (error) {
      if (!isReadonlyWalRetryError(error, dbPath)) throw error;
      blobs = queryCoverBlobs(dbPath, missing, { readwrite: true, create: false });
    }
  } catch {
    return resolved;
  }
  if (blobs.size === 0) return resolved;

  try {
    fs.mkdirSync(cacheDir, { recursive: true });
  } catch {
    return resolved;
  }
  for (const [hash, blob] of blobs) {
    const target = path.join(cacheDir, `${hash}${detectImageExtension(blob)}`);
    try {
      fs.writeFileSync(target, blob);
      resolved.set(hash, target);
    } catch {
      // cache write failure just means no icon for this entry
    }
  }
  return resolved;
}
