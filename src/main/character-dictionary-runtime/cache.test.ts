import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import { getSnapshotPath, readSnapshot, writeSnapshot } from './cache';
import { CHARACTER_DICTIONARY_FORMAT_VERSION } from './constants';
import type { CharacterDictionarySnapshot } from './types';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-character-dictionary-cache-'));
}

function createSnapshot(): CharacterDictionarySnapshot {
  return {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId: 130298,
    mediaTitle: 'The Eminence in Shadow',
    entryCount: 1,
    updatedAt: 1_700_000_000_000,
    termEntries: [['アルファ', 'あるふぁ', '', '', 0, ['Alpha'], 0, 'name']],
    images: [
      {
        path: 'img/m130298-c1.png',
        dataBase64:
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nmX8AAAAASUVORK5CYII=',
      },
    ],
  };
}

test('writeSnapshot persists and readSnapshot restores current-format snapshots', () => {
  const outputDir = makeTempDir();
  const snapshotPath = getSnapshotPath(outputDir, 130298);
  const snapshot = createSnapshot();

  writeSnapshot(snapshotPath, snapshot);

  assert.deepEqual(readSnapshot(snapshotPath), snapshot);
});

test('readSnapshot ignores snapshots written with an older format version', () => {
  const outputDir = makeTempDir();
  const snapshotPath = getSnapshotPath(outputDir, 130298);
  const staleSnapshot = {
    ...createSnapshot(),
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION - 1,
  };

  fs.mkdirSync(path.dirname(snapshotPath), { recursive: true });
  fs.writeFileSync(snapshotPath, JSON.stringify(staleSnapshot), 'utf8');

  assert.equal(readSnapshot(snapshotPath), null);
});
