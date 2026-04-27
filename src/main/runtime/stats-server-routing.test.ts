import assert from 'node:assert/strict';
import test from 'node:test';
import { createEnsureStatsServerUrlHandler } from './stats-server-routing';

function createHarness(options?: {
  state?: { pid: number; port: number; startedAtMs: number } | null;
  processAlive?: boolean;
  localServerStarted?: boolean;
}) {
  const calls: string[] = [];
  let localServerStarted = options?.localServerStarted ?? false;
  const handler = createEnsureStatsServerUrlHandler({
    currentPid: 100,
    readBackgroundState: () => {
      calls.push('readBackgroundState');
      return options?.state ?? null;
    },
    removeBackgroundState: () => {
      calls.push('removeBackgroundState');
    },
    isProcessAlive: () => {
      calls.push('isProcessAlive');
      return options?.processAlive ?? true;
    },
    hasLocalStatsServer: () => localServerStarted,
    startLocalStatsServer: () => {
      calls.push('startLocalStatsServer');
      localServerStarted = true;
    },
    getConfiguredPort: () => 6969,
  });

  return {
    calls,
    handler,
  };
}

test('stats server routing defers to a live background daemon from another process', () => {
  const { calls, handler } = createHarness({
    state: { pid: 200, port: 7979, startedAtMs: 1 },
    processAlive: true,
  });

  assert.equal(handler(), 'http://127.0.0.1:7979');
  assert.deepEqual(calls, ['readBackgroundState', 'isProcessAlive']);
});

test('stats server routing clears dead daemon state and starts local server', () => {
  const { calls, handler } = createHarness({
    state: { pid: 200, port: 7979, startedAtMs: 1 },
    processAlive: false,
  });

  assert.equal(handler(), 'http://127.0.0.1:6969');
  assert.deepEqual(calls, [
    'readBackgroundState',
    'isProcessAlive',
    'removeBackgroundState',
    'startLocalStatsServer',
  ]);
});

test('stats server routing clears self-owned stale state and starts local server', () => {
  const { calls, handler } = createHarness({
    state: { pid: 100, port: 7979, startedAtMs: 1 },
    processAlive: true,
  });

  assert.equal(handler(), 'http://127.0.0.1:6969');
  assert.deepEqual(calls, [
    'readBackgroundState',
    'removeBackgroundState',
    'startLocalStatsServer',
  ]);
});

test('stats server routing reuses a started local stats server', () => {
  const { calls, handler } = createHarness({
    state: null,
    localServerStarted: true,
  });

  assert.equal(handler(), 'http://127.0.0.1:6969');
  assert.deepEqual(calls, ['readBackgroundState', 'removeBackgroundState']);
});
