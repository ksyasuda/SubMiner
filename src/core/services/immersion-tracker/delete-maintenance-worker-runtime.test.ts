import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { DeleteMaintenanceWorkerRuntime } from './delete-maintenance-worker-runtime';
import { executeDeleteMaintenanceTask } from './delete-maintenance';
import { startSessionRecord } from './session';
import { Database } from './sqlite';
import { applyPragmas, ensureSchema, getOrCreateVideoRecord } from './storage';

test('a delete batch rebuilds lifetime summaries once', () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-delete-batch-test-'));
  const dbPath = path.join(tempDir, 'immersion.sqlite');
  let db = new Database(dbPath);

  try {
    applyPragmas(db);
    ensureSchema(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/batch-delete.mkv', {
      canonicalTitle: 'Batch Delete',
      sourcePath: '/tmp/batch-delete.mkv',
      sourceUrl: null,
      sourceType: 1,
    });
    const firstSessionId = startSessionRecord(db, videoId, 1_000).sessionId;
    const secondSessionId = startSessionRecord(db, videoId, 2_000).sessionId;
    const deletedVideoId = getOrCreateVideoRecord(db, 'local:/tmp/batch-delete-video.mkv', {
      canonicalTitle: 'Batch Delete Video',
      sourcePath: '/tmp/batch-delete-video.mkv',
      sourceUrl: null,
      sourceType: 1,
    });
    startSessionRecord(db, deletedVideoId, 3_000);
    db.exec(`
      CREATE TABLE delete_rebuild_audit (id INTEGER PRIMARY KEY);
      CREATE TRIGGER count_delete_lifetime_rebuild
      AFTER UPDATE OF last_rebuilt_ms ON imm_lifetime_global
      BEGIN
        INSERT INTO delete_rebuild_audit (id) VALUES (NULL);
      END;
    `);
    db.close();

    executeDeleteMaintenanceTask(dbPath, {
      kind: 'batch',
      tasks: [
        { kind: 'session', sessionId: firstSessionId },
        { kind: 'video', videoId: deletedVideoId },
      ],
    });

    db = new Database(dbPath);
    const audit = db.prepare('SELECT COUNT(*) AS total FROM delete_rebuild_audit').get() as {
      total: number;
    };
    const retainedSession = db
      .prepare('SELECT session_id AS sessionId FROM imm_sessions WHERE video_id = ?')
      .get(videoId) as { sessionId: number } | null;
    const deletedVideo = db
      .prepare('SELECT video_id AS videoId FROM imm_videos WHERE video_id = ?')
      .get(deletedVideoId) as { videoId: number } | null;
    assert.equal(retainedSession?.sessionId, secondSessionId);
    assert.equal(deletedVideo, undefined);
    assert.equal(
      audit.total,
      2,
      'one rebuild performs exactly its reset and final global summary writes',
    );
  } finally {
    try {
      db.close();
    } catch {
      // The setup connection closes before maintenance runs.
    }
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test(
  'compiled delete worker removes data through its separate database connection',
  { skip: path.extname(__filename) !== '.js' },
  async () => {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-delete-worker-test-'));
    const dbPath = path.join(tempDir, 'immersion.sqlite');
    const runtime = new DeleteMaintenanceWorkerRuntime();
    let db = new Database(dbPath);

    try {
      applyPragmas(db);
      ensureSchema(db);
      const videoId = getOrCreateVideoRecord(db, 'local:/tmp/worker-delete.mkv', {
        canonicalTitle: 'Worker Delete',
        sourcePath: '/tmp/worker-delete.mkv',
        sourceUrl: null,
        sourceType: 1,
      });
      const firstSessionId = startSessionRecord(db, videoId, 1_000).sessionId;
      const secondSessionId = startSessionRecord(db, videoId, 2_000).sessionId;
      db.close();

      await runtime.run(dbPath, {
        kind: 'batch',
        tasks: [
          { kind: 'session', sessionId: firstSessionId },
          { kind: 'session', sessionId: secondSessionId },
        ],
      });

      db = new Database(dbPath);
      const row = db
        .prepare('SELECT COUNT(*) AS total FROM imm_sessions WHERE video_id = ?')
        .get(videoId) as { total: number };
      assert.equal(row.total, 0);
    } finally {
      runtime.destroy();
      try {
        db.close();
      } catch {
        // The setup connection is already closed before the worker starts.
      }
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  },
);
