import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  LexicalRollupWorkerRuntime,
  resolveLexicalRollupWorkerPath,
} from './lexical-rollup-worker-runtime';
import { areLexicalDailyRollupsReady } from './lexical-rollups';
import { Database } from './sqlite';
import { applyPragmas, ensureSchema } from './storage';

test('lexical rollup worker backfills without using the tracker connection', async () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-lexical-rollup-runtime-'));
  const dbPath = path.join(directory, 'immersion.sqlite');
  const runtime = new LexicalRollupWorkerRuntime();
  const db = new Database(dbPath);

  try {
    applyPragmas(db);
    ensureSchema(db);
    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES ('鳥', '鳥', 'とり', 1700000000, 1700000000, 1)`,
    ).run();
    db.exec('DELETE FROM imm_lexical_daily_rollups');
    db.prepare(`UPDATE imm_rollup_state SET state_value = '0' WHERE state_key = ?`).run(
      'lexical_daily_rollups_version',
    );
    db.close();

    await runtime.run(dbPath);

    const checkDb = new Database(dbPath);
    try {
      assert.equal(areLexicalDailyRollupsReady(checkDb), true);
    } finally {
      checkDb.close();
    }
  } finally {
    runtime.destroy();
    try {
      db.close();
    } catch {
      // Closed before the worker starts.
    }
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

test('lexical rollup worker module resolves in the current layout', () => {
  const workerPath = resolveLexicalRollupWorkerPath();
  assert.ok(workerPath, 'expected the lexical rollup worker module to resolve');
  assert.ok(workerPath.endsWith(__filename.endsWith('.ts') ? '.ts' : '.js'));
});

test('lexical rollup worker leaves a backfill pending when no worker can start', async () => {
  const runtime = new LexicalRollupWorkerRuntime({
    resolveWorkerPath: () => null,
    warn: () => {},
  } as never);

  try {
    await assert.doesNotReject(runtime.run('/tmp/not-used.sqlite'));
  } finally {
    runtime.destroy();
  }
});

test('lexical rollup worker absorbs termination failures after settling', async () => {
  let sendMessage: ((message: { ok: boolean }) => void) | null = null;
  const runtime = new LexicalRollupWorkerRuntime({
    resolveWorkerPath: () => '/tmp/fake-worker.js',
    createWorker: async () => ({
      once(event: string, listener: (value: never) => void) {
        if (event === 'message') sendMessage = listener as (message: { ok: boolean }) => void;
        return this;
      },
      terminate: async () => {
        throw new Error('termination failed');
      },
    }),
    warn: () => {},
  } as never);

  const unhandled: unknown[] = [];
  const captureUnhandled = (reason: unknown) => unhandled.push(reason);
  process.on('unhandledRejection', captureUnhandled);
  try {
    const task = runtime.run('/tmp/not-used.sqlite');
    await new Promise((resolve) => setImmediate(resolve));
    const notify = sendMessage as ((message: { ok: boolean }) => void) | null;
    assert.ok(notify);
    notify({ ok: true });
    await task;
    await new Promise((resolve) => setImmediate(resolve));
    assert.deepEqual(unhandled, []);
  } finally {
    process.off('unhandledRejection', captureUnhandled);
    runtime.destroy();
  }
});

test('lexical rollup worker times out when it never responds', async () => {
  let terminated = false;
  const runtime = new LexicalRollupWorkerRuntime({
    resolveWorkerPath: () => '/tmp/fake-worker.js',
    createWorker: async () => ({
      once() {
        return this;
      },
      terminate: async () => {
        terminated = true;
        return 0;
      },
    }),
    timeoutMs: 1,
    warn: () => {},
  } as never);

  try {
    const outcome = await Promise.race([
      runtime.run('/tmp/not-used.sqlite').then(
        () => 'resolved',
        (error: unknown) => String(error),
      ),
      new Promise<string>((resolve) => setTimeout(() => resolve('still pending'), 50)),
    ]);

    assert.match(outcome, /timed out/);
    assert.equal(terminated, true);
  } finally {
    runtime.destroy();
  }
});

test('destroy waits for active worker termination', async () => {
  let releaseTermination = (): void => {};
  let emitExit: ((code: number) => void) | null = null;
  const terminationGate = new Promise<void>((resolve) => {
    releaseTermination = resolve;
  });
  const runtime = new LexicalRollupWorkerRuntime({
    resolveWorkerPath: () => '/tmp/fake-worker.js',
    createWorker: async () => ({
      once(event: string, listener: (value: never) => void) {
        if (event === 'exit') emitExit = listener as (code: number) => void;
        return this;
      },
      terminate: async () => {
        await terminationGate;
        emitExit?.(1);
        return 1;
      },
    }),
    warn: () => {},
  } as never);

  const runTask = runtime.run('/tmp/not-used.sqlite');
  await new Promise((resolve) => setImmediate(resolve));
  const destroyTask = runtime.destroy();
  assert.ok(destroyTask instanceof Promise);

  let destroyed = false;
  void destroyTask.then(() => {
    destroyed = true;
  });
  await Promise.resolve();
  assert.equal(destroyed, false);

  releaseTermination();
  await destroyTask;
  await assert.rejects(runTask, /exited with code 1/);
  assert.equal(destroyed, true);
});
