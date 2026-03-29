import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  loadSubtitlePosition,
  saveSubtitlePosition,
  updateCurrentMediaPath,
} from './subtitle-position';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subtitle-position-test-'));
}

test('saveSubtitlePosition queues pending position when media path is unavailable', () => {
  const queued: Array<{ yPercent: number }> = [];
  let persisted = false;

  saveSubtitlePosition({
    position: { yPercent: 21 },
    currentMediaPath: null,
    subtitlePositionsDir: makeTempDir(),
    onQueuePending: (position) => {
      queued.push(position);
    },
    onPersisted: () => {
      persisted = true;
    },
  });

  assert.deepEqual(queued, [{ yPercent: 21 }]);
  assert.equal(persisted, false);
});

test('saveSubtitlePosition persists and loadSubtitlePosition restores the stored position', () => {
  const dir = makeTempDir();
  const mediaPath = path.join(dir, 'episode.mkv');
  const position = { yPercent: 37 };
  let persisted = false;

  saveSubtitlePosition({
    position,
    currentMediaPath: mediaPath,
    subtitlePositionsDir: dir,
    onQueuePending: () => {
      throw new Error('unexpected queue');
    },
    onPersisted: () => {
      persisted = true;
    },
  });

  const loaded = loadSubtitlePosition({
    currentMediaPath: mediaPath,
    fallbackPosition: { yPercent: 0 },
    subtitlePositionsDir: dir,
  });

  assert.equal(persisted, true);
  assert.deepEqual(loaded, position);
  assert.equal(
    fs.readdirSync(dir).some((entry) => entry.endsWith('.json')),
    true,
  );
});

test('updateCurrentMediaPath persists a queued subtitle position before broadcasting', () => {
  const dir = makeTempDir();
  let currentMediaPath: string | null = null;
  let cleared = false;
  const setPositions: Array<{ yPercent: number } | null> = [];
  const broadcasts: Array<{ yPercent: number } | null> = [];
  const pending = { yPercent: 64 };

  updateCurrentMediaPath({
    mediaPath: path.join(dir, 'video.mkv'),
    currentMediaPath,
    pendingSubtitlePosition: pending,
    subtitlePositionsDir: dir,
    loadSubtitlePosition: () =>
      loadSubtitlePosition({
        currentMediaPath,
        fallbackPosition: { yPercent: 0 },
        subtitlePositionsDir: dir,
      }),
    setCurrentMediaPath: (next) => {
      currentMediaPath = next;
    },
    clearPendingSubtitlePosition: () => {
      cleared = true;
    },
    setSubtitlePosition: (position) => {
      setPositions.push(position);
    },
    broadcastSubtitlePosition: (position) => {
      broadcasts.push(position);
    },
  });

  assert.equal(currentMediaPath, path.join(dir, 'video.mkv'));
  assert.equal(cleared, true);
  assert.deepEqual(setPositions, [pending]);
  assert.deepEqual(broadcasts, [pending]);
  assert.deepEqual(
    loadSubtitlePosition({
      currentMediaPath,
      fallbackPosition: { yPercent: 0 },
      subtitlePositionsDir: dir,
    }),
    pending,
  );
});
