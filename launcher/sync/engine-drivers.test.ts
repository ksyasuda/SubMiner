import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  createImmersionDbFixture,
  insertFixtureSession,
} from '../test-support/immersion-db-fixture.js';
import { openLibsqlSyncDb } from '../../src/core/services/stats-sync/libsql-driver.js';
import { createDbSnapshot } from '../../src/core/services/stats-sync/shared.js';
import { mergeSnapshotIntoDb } from '../../src/core/services/stats-sync/merge.js';

const BASE_MS = Date.UTC(2026, 5, 1, 12, 0, 0);

function makeTmpDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-libsql-sync-test-'));
}

function countRows(dbPath: string, sql: string): number {
  const db = openLibsqlSyncDb(dbPath, { readonly: true });
  try {
    const row = db.query(sql).get() as { n: number } | undefined;
    return Number(row?.n ?? 0);
  } finally {
    db.close();
  }
}

// End-to-end pass of the shared engine over the libsql driver (the binding the
// Electron app's --sync-cli mode uses): snapshot a fixture DB, merge it into
// another, and re-merge to confirm idempotence.
test('libsql driver runs snapshot and merge through the shared engine', () => {
  const dir = makeTmpDir();
  try {
    const localPath = path.join(dir, 'local.sqlite');
    const remotePath = path.join(dir, 'remote.sqlite');
    createImmersionDbFixture(localPath);
    createImmersionDbFixture(remotePath);
    insertFixtureSession(localPath, {
      uuid: 'local-1',
      videoKey: 'showa-e1',
      animeTitleKey: 'showa',
      startedAtMs: BASE_MS,
      applyLifetime: true,
      words: [{ headword: '見る', word: '見た', reading: 'みた', count: 2 }],
    });
    insertFixtureSession(remotePath, {
      uuid: 'remote-1',
      videoKey: 'showb-e1',
      animeTitleKey: 'showb',
      startedAtMs: BASE_MS + 86_400_000,
      activeWatchedMs: 900_000,
      cardsMined: 3,
      applyLifetime: true,
      words: [{ headword: '食べる', word: '食べた', reading: 'たべた', count: 1 }],
    });

    const snapshotPath = path.join(dir, 'remote-snapshot.sqlite');
    createDbSnapshot(openLibsqlSyncDb, remotePath, snapshotPath);
    assert.ok(fs.existsSync(snapshotPath));
    assert.equal(countRows(snapshotPath, 'SELECT COUNT(*) AS n FROM imm_sessions'), 1);

    const summary = mergeSnapshotIntoDb(openLibsqlSyncDb, localPath, snapshotPath);
    assert.equal(summary.sessionsMerged, 1);
    assert.equal(summary.videosAdded, 1);
    assert.equal(countRows(localPath, 'SELECT COUNT(*) AS n FROM imm_sessions'), 2);

    const again = mergeSnapshotIntoDb(openLibsqlSyncDb, localPath, snapshotPath);
    assert.equal(again.sessionsMerged, 0);
    assert.equal(again.sessionsAlreadyPresent, 1);
    assert.equal(countRows(localPath, 'SELECT COUNT(*) AS n FROM imm_sessions'), 2);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
