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
  while (releases.length === 0) await new Promise<void>((resolve) => setTimeout(resolve, 0));
  const queued = scheduler.enqueue(() => ({ kind: 'session', sessionId: 2 }));
  scheduler.destroy();

  await assert.rejects(queued, /shutting down/);
  releases[0]?.();
  await first;
  assert.equal(maxActiveTasks, 1);
});
