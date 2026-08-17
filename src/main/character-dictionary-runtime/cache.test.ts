import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import { isDeepStrictEqual } from 'node:util';
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
  // Distinct titles, lengths, and term text so the surviving file can be pinned to exactly one
  // writer rather than merely "a snapshot that parses". A shared temp file is caught by the
  // losing writers failing to rename; interleaved content is only caught when the timing happens
  // to leave a mix, which is why the assertion checks identity rather than shape.
  const variants: CharacterDictionarySnapshot[] = ['alpha', 'beta', 'gamma'].map((label, index) => {
    const entryCount = 400 + index * 100;
    return {
      ...base,
      mediaTitle: `${base.mediaTitle} ${label}`,
      entryCount,
      termEntries: Array.from({ length: entryCount }, (_entry, entryIndex) => [
        `${label}${entryIndex}`,
        'なまえ',
        'name primary',
        '',
        75,
        [`${label} character ${entryIndex} `.repeat(600)],
        0,
        '',
      ]) as CharacterDictionarySnapshot['termEntries'],
    };
  });

  await Promise.all(variants.map((variant) => writeSnapshot(snapshotPath, variant)));

  const restored = await readSnapshot(snapshotPath);
  const expected = variants.map((variant) => ({
    ...variant,
    nameSplitSource: 'heuristic' as const,
  }));
  const matches = expected.filter((candidate) => isDeepStrictEqual(restored, candidate));
  assert.equal(
    matches.length,
    1,
    `expected exactly one writer's complete snapshot to survive, got ${
      restored === null
        ? 'an unreadable file'
        : `entryCount=${restored.entryCount}, terms=${restored.termEntries.length}, title=${restored.mediaTitle}`
    }`,
  );

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
