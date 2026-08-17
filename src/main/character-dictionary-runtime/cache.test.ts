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

test('writeSnapshot persists and readSnapshot restores current-format snapshots', async () => {
  const outputDir = makeTempDir();
  const snapshotPath = getSnapshotPath(outputDir, 130298);
  const snapshot = createSnapshot();

  await writeSnapshot(snapshotPath, snapshot);

  assert.deepEqual(await readSnapshot(snapshotPath), { ...snapshot, nameSplitSource: 'heuristic' });
});

// A manual generate and an auto-sync can both land on the same media, so two writes for one
// snapshot can overlap. They must not stream into a shared temp file and interleave into a
// half-and-half snapshot.
test('concurrent writeSnapshot calls for the same media leave one complete snapshot', async () => {
  const outputDir = makeTempDir();
  const snapshotPath = getSnapshotPath(outputDir, 130298);
  const base = createSnapshot();
  const wide: CharacterDictionarySnapshot = {
    ...base,
    entryCount: 400,
    termEntries: Array.from({ length: 400 }, (_entry, index) => [
      `名前${index}`,
      'なまえ',
      'name primary',
      '',
      75,
      [`Character ${index}`.repeat(200)],
      0,
      '',
    ]) as CharacterDictionarySnapshot['termEntries'],
  };

  await Promise.all([
    writeSnapshot(snapshotPath, wide),
    writeSnapshot(snapshotPath, wide),
    writeSnapshot(snapshotPath, wide),
  ]);

  const restored = await readSnapshot(snapshotPath);
  assert.equal(restored?.entryCount, 400);
  assert.equal(restored?.termEntries.length, 400);
  assert.deepEqual(restored, { ...wide, nameSplitSource: 'heuristic' });

  // Every writer cleaned up after itself, so no temp files are left behind.
  const leftovers = fs
    .readdirSync(path.dirname(snapshotPath))
    .filter((name) => name.includes('.tmp-'));
  assert.deepEqual(leftovers, []);
});

test('readSnapshot preserves the mecab name-split source and defaults missing values to heuristic', async () => {
  const outputDir = makeTempDir();
  const snapshotPath = getSnapshotPath(outputDir, 130298);
  const snapshot: CharacterDictionarySnapshot = {
    ...createSnapshot(),
    nameSplitSource: 'mecab',
  };

  await writeSnapshot(snapshotPath, snapshot);

  assert.equal((await readSnapshot(snapshotPath))?.nameSplitSource, 'mecab');
});

test('readSnapshot ignores snapshots written with an older format version', async () => {
  const outputDir = makeTempDir();
  const snapshotPath = getSnapshotPath(outputDir, 130298);
  const staleSnapshot = {
    ...createSnapshot(),
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION - 1,
  };

  fs.mkdirSync(path.dirname(snapshotPath), { recursive: true });
  fs.writeFileSync(snapshotPath, JSON.stringify(staleSnapshot), 'utf8');

  assert.equal(await readSnapshot(snapshotPath), null);
});

test('readSnapshot ignores v15 snapshots with stale romanized character-name entries', async () => {
  const outputDir = makeTempDir();
  const snapshotPath = getSnapshotPath(outputDir, 130298);
  const staleSnapshot = {
    ...createSnapshot(),
    formatVersion: 15,
    termEntries: [['Vanir', 'ばにる', 'name primary', '', 75, ['Vanir'], 0, '']],
  };

  fs.mkdirSync(path.dirname(snapshotPath), { recursive: true });
  fs.writeFileSync(snapshotPath, JSON.stringify(staleSnapshot), 'utf8');

  assert.equal(await readSnapshot(snapshotPath), null);
});
