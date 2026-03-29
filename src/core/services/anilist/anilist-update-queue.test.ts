import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import { createAnilistUpdateQueue } from './anilist-update-queue';

function createTempQueueFile(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-anilist-queue-'));
  return path.join(dir, 'queue.json');
}

function createLogger() {
  const info: string[] = [];
  const warn: string[] = [];
  const error: string[] = [];
  return {
    info,
    warn,
    error,
    logger: {
      info: (message: string) => info.push(message),
      warn: (message: string) => warn.push(message),
      error: (message: string) => error.push(message),
    },
  };
}

test('anilist update queue enqueues, snapshots, and dequeues success', () => {
  const queueFile = createTempQueueFile();
  const loggerState = createLogger();
  const queue = createAnilistUpdateQueue(queueFile, loggerState.logger);

  queue.enqueue('k1', 'Demo', 1);
  const snapshot = queue.getSnapshot(Number.MAX_SAFE_INTEGER);
  assert.deepEqual(snapshot, { pending: 1, ready: 1, deadLetter: 0 });
  assert.equal(queue.nextReady(Number.MAX_SAFE_INTEGER)?.key, 'k1');

  queue.markSuccess('k1');
  assert.deepEqual(queue.getSnapshot(Number.MAX_SAFE_INTEGER), {
    pending: 0,
    ready: 0,
    deadLetter: 0,
  });
  assert.ok(loggerState.info.some((message) => message.includes('Queued AniList retry')));
});

test('anilist update queue applies retry backoff and dead-letter', () => {
  const queueFile = createTempQueueFile();
  const loggerState = createLogger();
  const queue = createAnilistUpdateQueue(queueFile, loggerState.logger);

  const now = 1_700_000 * 1_000_000;
  queue.enqueue('k2', 'Backoff Demo', 2);

  queue.markFailure('k2', 'fail-1', now);
  const firstRetry = queue.nextReady(now);
  assert.equal(firstRetry, null);

  const pendingPayload = JSON.parse(fs.readFileSync(queueFile, 'utf-8')) as {
    pending: Array<{ attemptCount: number; nextAttemptAt: number }>;
  };
  assert.equal(pendingPayload.pending[0]?.attemptCount, 1);
  assert.equal((pendingPayload.pending[0]?.nextAttemptAt ?? now) - now, 30_000);

  for (let attempt = 2; attempt <= 8; attempt += 1) {
    queue.markFailure('k2', `fail-${attempt}`, now);
  }

  const snapshot = queue.getSnapshot(Number.MAX_SAFE_INTEGER);
  assert.deepEqual(snapshot, { pending: 0, ready: 0, deadLetter: 1 });
  assert.ok(
    loggerState.warn.some((message) =>
      message.includes('AniList retry moved to dead-letter queue.'),
    ),
  );
});

test('anilist update queue persists and reloads from disk', () => {
  const queueFile = createTempQueueFile();
  const loggerState = createLogger();
  const queueA = createAnilistUpdateQueue(queueFile, loggerState.logger);
  queueA.enqueue('k3', 'Persist Demo', 3);

  const queueB = createAnilistUpdateQueue(queueFile, loggerState.logger);
  assert.deepEqual(queueB.getSnapshot(Number.MAX_SAFE_INTEGER), {
    pending: 1,
    ready: 1,
    deadLetter: 0,
  });
  assert.equal(queueB.nextReady(Number.MAX_SAFE_INTEGER)?.title, 'Persist Demo');
});
