import assert from 'node:assert/strict';
import test from 'node:test';
import { DeleteMaintenanceScheduler } from './delete-maintenance-scheduler';
import type { DeleteMaintenanceTask } from './delete-maintenance';

test('scheduler batches same-turn requests and balances busy state', async () => {
  const tasks: DeleteMaintenanceTask[] = [];
  const states: string[] = [];
  const scheduler = new DeleteMaintenanceScheduler({
    batchWindowMs: 0,
    runTask: async (task) => {
      tasks.push(task);
    },
    onBusy: () => states.push('busy'),
    onIdle: () => states.push('idle'),
  });

  const first = scheduler.enqueue(() => ({ kind: 'session', sessionId: 1 }));
  const second = scheduler.enqueue(() => ({ kind: 'sessions', sessionIds: [2, 3] }));
  const third = scheduler.enqueue(() => null);
  await Promise.all([first, second, third]);

  assert.deepEqual(tasks, [
    {
      kind: 'batch',
      tasks: [
        { kind: 'session', sessionId: 1 },
        { kind: 'sessions', sessionIds: [2, 3] },
      ],
    },
  ]);
  assert.deepEqual(states, ['busy', 'idle']);
});

test('scheduler rejects enqueue after destruction without entering busy state', async () => {
  let busyCalls = 0;
  let runCalls = 0;
  const scheduler = new DeleteMaintenanceScheduler({
    batchWindowMs: 0,
    runTask: async () => {
      runCalls += 1;
    },
    onBusy: () => {
      busyCalls += 1;
    },
    onIdle: () => {},
  });
  scheduler.destroy();

  await assert.rejects(
    scheduler.enqueue(() => ({ kind: 'session', sessionId: 1 })),
    /shutting down/,
  );
  assert.equal(busyCalls, 0);
  assert.equal(runCalls, 0);
});

test('scheduler rejects every request in a batch when the maintenance task fails', async () => {
  const failure = new Error('maintenance failed');
  const scheduler = new DeleteMaintenanceScheduler({
    batchWindowMs: 0,
    runTask: async () => {
      throw failure;
    },
    onBusy: () => {},
    onIdle: () => {},
  });

  const first = scheduler.enqueue(() => ({ kind: 'session', sessionId: 1 }));
  const second = scheduler.enqueue(() => ({ kind: 'session', sessionId: 2 }));

  const results = await Promise.allSettled([first, second]);
  assert.deepEqual(
    results.map((result) => (result.status === 'rejected' ? result.reason : null)),
    [failure, failure],
  );
});

test('scheduler rejects only the request whose task resolution fails', async () => {
  const failure = new Error('resolution failed');
  const tasks: DeleteMaintenanceTask[] = [];
  const scheduler = new DeleteMaintenanceScheduler({
    batchWindowMs: 0,
    runTask: async (task) => {
      tasks.push(task);
    },
    onBusy: () => {},
    onIdle: () => {},
  });

  const failed = scheduler.enqueue(() => {
    throw failure;
  });
  const succeeded = scheduler.enqueue(() => ({ kind: 'session', sessionId: 2 }));

  const results = await Promise.allSettled([failed, succeeded]);
  assert.equal(results[0]?.status, 'rejected');
  assert.equal(results[0]?.status === 'rejected' ? results[0].reason : null, failure);
  assert.equal(results[1]?.status, 'fulfilled');
  assert.deepEqual(tasks, [{ kind: 'session', sessionId: 2 }]);
});

test('scheduler does not schedule another drain when the queue is empty', async () => {
  const originalSetTimeout = globalThis.setTimeout;
  let timerCalls = 0;
  globalThis.setTimeout = ((handler: TimerHandler, timeout?: number, ...args: unknown[]) => {
    timerCalls += 1;
    return originalSetTimeout(handler, timeout, ...args);
  }) as typeof setTimeout;

  try {
    const scheduler = new DeleteMaintenanceScheduler({
      batchWindowMs: 0,
      runTask: async () => {},
      onBusy: () => {},
      onIdle: () => {},
    });

    await scheduler.enqueue(() => ({ kind: 'session', sessionId: 1 }));
    assert.equal(timerCalls, 1);
  } finally {
    globalThis.setTimeout = originalSetTimeout;
  }
});

test('scheduler serializes batches and rejects requests queued at destruction', async () => {
  const releases: Array<() => void> = [];
  let activeTasks = 0;
  let maxActiveTasks = 0;
  const scheduler = new DeleteMaintenanceScheduler({
    batchWindowMs: 0,
    runTask: async () => {
      activeTasks += 1;
      maxActiveTasks = Math.max(maxActiveTasks, activeTasks);
      await new Promise<void>((resolve) => releases.push(resolve));
      activeTasks -= 1;
    },
    onBusy: () => {},
    onIdle: () => {},
  });

  const first = scheduler.enqueue(() => ({ kind: 'session', sessionId: 1 }));
  const maxPollAttempts = 100;
  let pollAttempts = 0;
  while (releases.length === 0 && pollAttempts < maxPollAttempts) {
    pollAttempts += 1;
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  }
  assert.ok(
    releases.length > 0,
    `runTask did not produce a release after ${maxPollAttempts} polling attempts`,
  );
  const queued = scheduler.enqueue(() => ({ kind: 'session', sessionId: 2 }));
  scheduler.destroy();

  await assert.rejects(queued, /shutting down/);
  releases[0]?.();
  await first;
  assert.equal(maxActiveTasks, 1);
});
