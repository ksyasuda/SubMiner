import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  DeleteMaintenanceWorkerRuntime,
  resolveDeleteMaintenanceWorkerPath,
} from './delete-maintenance-worker-runtime';
import { executeDeleteMaintenanceTask } from './delete-maintenance';
import { startSessionRecord } from './session';
import { Database } from './sqlite';
import { applyPragmas, ensureSchema, getOrCreateVideoRecord } from './storage';

type FakeWorkerListener = (value: never) => void;

function createFakeWorker() {
  const listeners = new Map<string, FakeWorkerListener>();
  const terminationState = { calls: 0 };
  const worker = {
    once(event: string, listener: FakeWorkerListener) {
      listeners.set(event, listener);
      return this;
    },
    terminate: async () => {
      terminationState.calls += 1;
      return 0;
    },
  };
  return { worker, listeners, terminationState };
}

type FakeWorker = ReturnType<typeof createFakeWorker>['worker'];

test('a delete batch never runs a full lifetime rebuild', () => {
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
      0,
      'delete maintenance subtracts incrementally instead of rewriting last_rebuilt_ms',
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

test('delete worker module resolves in the current layout', () => {
  // If this resolves to null, every delete silently runs on the serving thread
  // and blocks the stats API for the whole maintenance run.
  const workerPath = resolveDeleteMaintenanceWorkerPath();
  assert.ok(workerPath, 'delete-maintenance worker module must resolve');
  assert.ok(workerPath.endsWith(__filename.endsWith('.ts') ? '.ts' : '.js'));
});

test(
  'compiled delete worker removes data through its separate database connection',
  { skip: resolveDeleteMaintenanceWorkerPath() === null },
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

test('worker runtime warns before falling back when no emitted worker is available', async () => {
  const warnings: unknown[][] = [];
  const fallbackTasks: unknown[] = [];
  const runtime = new DeleteMaintenanceWorkerRuntime({
    resolveWorkerPath: () => null,
    warn: (...args) => warnings.push(args),
    executeFallback: (_dbPath, task) => fallbackTasks.push(task),
  });

  await runtime.run('/tmp/fallback.sqlite', { kind: 'session', sessionId: 1 });

  assert.equal(warnings.length, 1);
  assert.match(String(warnings[0]?.[0]), /worker unavailable/i);
  assert.deepEqual(fallbackTasks, [{ kind: 'session', sessionId: 1 }]);
});

test('worker runtime terminates a worker after successful settlement', async () => {
  const { worker, listeners, terminationState } = createFakeWorker();
  const runtime = new DeleteMaintenanceWorkerRuntime({
    resolveWorkerPath: () => '/tmp/delete-worker.js',
    createWorker: async () => worker,
  });

  const result = runtime.run('/tmp/test.sqlite', { kind: 'session', sessionId: 1 });
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  listeners.get('message')?.({ ok: true } as never);
  await result;

  assert.equal(terminationState.calls, 1);
});

test('worker runtime falls back to the current thread when the worker crashes', async () => {
  const { worker, listeners, terminationState } = createFakeWorker();
  const fallbackTasks: unknown[] = [];
  const warnings: string[] = [];
  const runtime = new DeleteMaintenanceWorkerRuntime({
    resolveWorkerPath: () => '/tmp/delete-worker.js',
    createWorker: async () => worker,
    executeFallback: (dbPath, task) => {
      fallbackTasks.push({ dbPath, task });
    },
    warn: (message) => {
      warnings.push(message);
    },
  });

  const result = runtime.run('/tmp/test.sqlite', { kind: 'session', sessionId: 1 });
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  listeners.get('error')?.(new Error('worker failed') as never);

  await result;
  assert.equal(terminationState.calls, 1);
  assert.deepEqual(fallbackTasks, [
    { dbPath: '/tmp/test.sqlite', task: { kind: 'session', sessionId: 1 } },
  ]);
  assert.equal(warnings.length, 1);
});

test('worker runtime surfaces a task failure without rerunning it', async () => {
  const { worker, listeners, terminationState } = createFakeWorker();
  const fallbackTasks: unknown[] = [];
  const runtime = new DeleteMaintenanceWorkerRuntime({
    resolveWorkerPath: () => '/tmp/delete-worker.js',
    createWorker: async () => worker,
    executeFallback: () => {
      fallbackTasks.push('ran');
    },
  });

  const result = runtime.run('/tmp/test.sqlite', { kind: 'session', sessionId: 1 });
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  listeners.get('message')?.({ ok: false, error: 'constraint violated' } as never);

  await assert.rejects(result, /constraint violated/);
  assert.equal(terminationState.calls, 1);
  assert.equal(fallbackTasks.length, 0);
});

test('worker runtime terminates a worker created after shutdown begins', async () => {
  const { worker, listeners, terminationState } = createFakeWorker();
  const createGate: { resolve?: (worker: FakeWorker) => void } = {};
  const fallbackTasks: unknown[] = [];
  const runtime = new DeleteMaintenanceWorkerRuntime({
    resolveWorkerPath: () => '/tmp/delete-worker.js',
    createWorker: () =>
      new Promise((resolve) => {
        createGate.resolve = resolve;
      }),
    executeFallback: (_dbPath, task) => fallbackTasks.push(task),
  });

  const result = runtime.run('/tmp/test.sqlite', { kind: 'session', sessionId: 1 });
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  runtime.destroy();
  createGate.resolve?.(worker);

  await assert.rejects(result, /shut down/);
  assert.equal(terminationState.calls, 1);
  assert.equal(listeners.size, 0);
  assert.deepEqual(fallbackTasks, []);
});

test('worker runtime does not fall back when worker creation fails during shutdown', async () => {
  const createGate: { reject?: (error: Error) => void } = {};
  const fallbackTasks: unknown[] = [];
  const runtime = new DeleteMaintenanceWorkerRuntime({
    resolveWorkerPath: () => '/tmp/delete-worker.js',
    createWorker: () =>
      new Promise((_resolve, reject) => {
        createGate.reject = reject;
      }),
    executeFallback: (_dbPath, task) => fallbackTasks.push(task),
  });

  const result = runtime.run('/tmp/test.sqlite', { kind: 'session', sessionId: 1 });
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  runtime.destroy();
  createGate.reject?.(new Error('creation failed'));

  await assert.rejects(result, /shut down/);
  assert.deepEqual(fallbackTasks, []);
});
