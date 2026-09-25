import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { getDefaultConfigDir } from '../../../shared/setup-state';

export function transferCacheKey(peer: string): string {
  return createHash('sha256').update(peer).digest('hex');
}

export function isTransferCacheKey(value: string): boolean {
  return /^[a-f0-9]{64}$/.test(value);
}

/**
 * Keep one previously received snapshot per peer as an rsync basis. Copies
 * isolate active transfers from concurrent cache replacements. A missing or
 * unusable cache only costs bandwidth; it must never prevent a sync.
 */
export function createTransferCache(
  directory = path.join(getDefaultConfigDir(), 'sync-transfer-cache'),
) {
  function cachePath(key: string): string {
    if (!isTransferCacheKey(key)) throw new Error('Invalid sync transfer cache key');
    return path.join(directory, `${key}.sqlite`);
  }

  return {
    seed(key: string, tempDir: string): void {
      const source = cachePath(key);
      const incoming = path.join(tempDir, 'incoming', 'snapshot.sqlite');
      try {
        fs.mkdirSync(path.dirname(incoming), { recursive: true, mode: 0o700 });
        fs.copyFileSync(source, incoming, fs.constants.COPYFILE_FICLONE);
      } catch {
        // A cold transfer sends a complete compressed snapshot.
      }
    },

    remember(key: string, tempDir: string): void {
      const target = cachePath(key);
      const incoming = path.join(tempDir, 'incoming', 'snapshot.sqlite');
      let staging = '';
      try {
        if (!fs.existsSync(incoming)) return;
        fs.mkdirSync(directory, { recursive: true, mode: 0o700 });
        staging = fs.mkdtempSync(path.join(directory, '.write-'));
        const snapshot = path.join(staging, 'snapshot.sqlite');
        fs.copyFileSync(incoming, snapshot, fs.constants.COPYFILE_FICLONE);
        fs.renameSync(snapshot, target);
      } catch {
        // An older basis is still valid. Never publish a partially copied file.
      } finally {
        try {
          if (staging) fs.rmSync(staging, { recursive: true, force: true });
        } catch {
          // Cache cleanup is optional too.
        }
      }
    },
  };
}
