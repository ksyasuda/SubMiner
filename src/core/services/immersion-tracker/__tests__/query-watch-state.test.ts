import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Database } from '../sqlite.js';
import { ensureSchema, getOrCreateVideoRecord } from '../storage.js';
import { startSessionRecord } from '../session.js';
import { markVideoWatched } from '../query-maintenance.js';
import { getWatchStateByVideoKeys } from '../query-watch-state.js';
import { SOURCE_TYPE_REMOTE } from '../types.js';

function createDb() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-imm-watch-state-test-'));
  const dbPath = path.join(dir, 'immersion.sqlite');
  const db = new Database(dbPath);
  ensureSchema(db);
  return { db, dir };
}

function addStreamVideo(db: ReturnType<typeof createDb>['db'], statsPath: string): number {
  return getOrCreateVideoRecord(db, `remote:${statsPath}`, {
    canonicalTitle: statsPath,
    sourcePath: null,
    sourceUrl: statsPath,
    sourceType: SOURCE_TYPE_REMOTE,
  });
}

test('getWatchStateByVideoKeys reports watched marks and the newest session', () => {
  const { db, dir } = createDb();
  try {
    const watchedPath = 'animebrowser://src/anime/ep1';
    const startedPath = 'animebrowser://src/anime/ep2';
    const watchedId = addStreamVideo(db, watchedPath);
    const startedId = addStreamVideo(db, startedPath);

    startSessionRecord(db, watchedId, 1_000_000);
    startSessionRecord(db, watchedId, 3_000_000);
    startSessionRecord(db, startedId, 2_000_000);
    markVideoWatched(db, watchedId, true);

    const rows = getWatchStateByVideoKeys(db, [
      `remote:${watchedPath}`,
      `remote:${startedPath}`,
      'remote:animebrowser://src/anime/never-played',
    ]);
    const byKey = new Map(rows.map((row) => [row.videoKey, row]));

    assert.equal(rows.length, 2, 'a key with no video row comes back absent, not unwatched');
    assert.deepEqual(byKey.get(`remote:${watchedPath}`), {
      videoKey: `remote:${watchedPath}`,
      watched: true,
      lastWatchedMs: 3_000_000,
      sessionCount: 2,
    });
    // Started but never finished: a row exists, the watch mark does not.
    assert.equal(byKey.get(`remote:${startedPath}`)?.watched, false);
    assert.equal(byKey.get(`remote:${startedPath}`)?.lastWatchedMs, 2_000_000);
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('getWatchStateByVideoKeys handles a video that was never played', () => {
  const { db, dir } = createDb();
  try {
    const statsPath = 'animebrowser://src/anime/ep3';
    addStreamVideo(db, statsPath);
    const [row] = getWatchStateByVideoKeys(db, [`remote:${statsPath}`]);
    assert.equal(row?.lastWatchedMs, null);
    assert.equal(row?.sessionCount, 0);
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('getWatchStateByVideoKeys ignores empty keys and dedupes the rest', () => {
  const { db, dir } = createDb();
  try {
    const statsPath = 'animebrowser://src/anime/ep4';
    addStreamVideo(db, statsPath);
    const rows = getWatchStateByVideoKeys(db, ['', `remote:${statsPath}`, `remote:${statsPath}`]);
    assert.equal(rows.length, 1);
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
