import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createDefaultSyncHostsState,
  upsertSyncHost,
  recordSyncResult,
  type SyncHostsState,
} from '../../shared/sync/sync-hosts-store';
import { createSyncAutoScheduler } from './sync-auto-scheduler';

function makeState(): SyncHostsState {
  let state = createDefaultSyncHostsState();
  state = upsertSyncHost(state, { host: 'auto-box', autoSync: true, direction: 'both' }, 0);
  state = upsertSyncHost(state, { host: 'manual-box', autoSync: false }, 0);
  return state;
}

test('tick triggers a due auto-sync host with its saved direction', () => {
  let state = makeState();
  state = upsertSyncHost(state, { host: 'auto-box', autoSync: true, direction: 'pull' }, 0);
  const triggered: Array<{ host: string; direction: string }> = [];
  const scheduler = createSyncAutoScheduler({
    readState: () => state,
    isRunning: () => false,
    triggerHostSync: (host, direction) => triggered.push({ host, direction }),
    nowMs: () => 100 * 60_000,
  });

  scheduler.tick();
  assert.deepEqual(triggered, [{ host: 'auto-box', direction: 'pull' }]);
});

test('tick does not require the resident app and stats writer to stop', () => {
  const state = makeState();
  const triggered: string[] = [];
  const scheduler = createSyncAutoScheduler({
    readState: () => state,
    isRunning: () => false,
    triggerHostSync: (host) => triggered.push(host),
    nowMs: () => 100 * 60_000,
  });

  assert.doesNotThrow(() => scheduler.tick());
  assert.deepEqual(triggered, ['auto-box']);
});

test('tick skips hosts synced more recently than the interval', () => {
  let state = makeState();
  state = recordSyncResult(state, 'auto-box', {
    atMs: 90 * 60_000,
    status: 'success',
    detail: null,
  });
  const triggered: string[] = [];
  const scheduler = createSyncAutoScheduler({
    readState: () => state,
    isRunning: () => false,
    triggerHostSync: (host) => triggered.push(host),
    nowMs: () => 100 * 60_000, // 10 minutes later, default interval 30
  });

  scheduler.tick();
  assert.deepEqual(triggered, []);
});

test('tick does nothing while a run is active', () => {
  const state = makeState();
  const triggered: string[] = [];
  const running = createSyncAutoScheduler({
    readState: () => state,
    isRunning: () => true,
    triggerHostSync: (host) => triggered.push(host),
    nowMs: () => 100 * 60_000,
  });
  running.tick();

  assert.deepEqual(triggered, []);
});

test('tick triggers at most one host per tick', () => {
  let state = makeState();
  state = upsertSyncHost(state, { host: 'second-box', autoSync: true }, 0);
  const triggered: string[] = [];
  const scheduler = createSyncAutoScheduler({
    readState: () => state,
    isRunning: () => false,
    triggerHostSync: (host) => triggered.push(host),
    nowMs: () => 100 * 60_000,
  });

  scheduler.tick();
  assert.equal(triggered.length, 1);
});

test('tick logs a synchronous trigger failure and remains usable', () => {
  const state = makeState();
  const logs: string[] = [];
  let attempts = 0;
  const scheduler = createSyncAutoScheduler({
    readState: () => state,
    isRunning: () => false,
    triggerHostSync: () => {
      attempts += 1;
      throw new Error('spawn exploded');
    },
    nowMs: () => 100 * 60_000,
    log: (message) => logs.push(message),
  });

  assert.doesNotThrow(() => scheduler.tick());
  assert.doesNotThrow(() => scheduler.tick());
  assert.equal(attempts, 2);
  assert.ok(logs.some((message) => message.includes('spawn exploded')));
});

test('start/stop manage the interval timer', () => {
  const state = makeState();
  let setCalls = 0;
  let clearCalls = 0;
  const scheduler = createSyncAutoScheduler({
    readState: () => state,
    isRunning: () => false,
    triggerHostSync: () => {},
    nowMs: () => 0,
    setIntervalFn: ((): NodeJS.Timeout => {
      setCalls += 1;
      return 42 as unknown as NodeJS.Timeout;
    }) as unknown as typeof setInterval,
    clearIntervalFn: (() => {
      clearCalls += 1;
    }) as unknown as typeof clearInterval,
  });

  scheduler.start();
  scheduler.start(); // idempotent
  assert.equal(setCalls, 1);
  scheduler.stop();
  assert.equal(clearCalls, 1);
});
