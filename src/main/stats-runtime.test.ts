import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { createStatsRuntime } from './stats-runtime';

function withTempDir(fn: (dir: string) => Promise<void> | void): Promise<void> | void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-stats-runtime-test-'));
  const result = fn(dir);
  if (result instanceof Promise) {
    return result.finally(() => {
      fs.rmSync(dir, { recursive: true, force: true });
    });
  }
  fs.rmSync(dir, { recursive: true, force: true });
}

test('stats runtime removes stale daemon state', async () => {
  await withTempDir(async (dir) => {
    const statePath = path.join(dir, 'stats-daemon.json');
    fs.writeFileSync(
      statePath,
      JSON.stringify({ pid: 99999, port: 6969, startedAtMs: 1_234 }, null, 2),
    );

    const runtime = createStatsRuntime({
      statsDaemonStatePath: statePath,
      getResolvedConfig: () => ({
        immersionTracking: { enabled: true },
        stats: { serverPort: 6969 },
      }),
      getImmersionTracker: () => ({}) as never,
      ensureImmersionTrackerStartedCore: () => {},
      startStatsServer: () => ({ close: () => {} }),
      openExternal: async () => {},
      exitAppWithCode: () => {},
      logInfo: () => {},
      logWarn: () => {},
      logError: () => {},
      getCurrentPid: () => 123,
      isProcessAlive: () => false,
    });

    assert.equal(runtime.readLiveBackgroundStatsDaemonState(), null);
    assert.equal(fs.existsSync(statePath), false);
  });
});

test('stats runtime starts background server and writes owned daemon state', async () => {
  await withTempDir(async (dir) => {
    const statePath = path.join(dir, 'stats-daemon.json');
    let startedPort: number | null = null;

    const runtime = createStatsRuntime({
      statsDaemonStatePath: statePath,
      getResolvedConfig: () => ({
        immersionTracking: { enabled: true },
        stats: { serverPort: 6970 },
      }),
      getImmersionTracker: () => ({}) as never,
      ensureImmersionTrackerStartedCore: () => {},
      startStatsServer: (port) => {
        startedPort = port;
        return { close: () => {} };
      },
      openExternal: async () => {},
      exitAppWithCode: () => {},
      logInfo: () => {},
      logWarn: () => {},
      logError: () => {},
      getCurrentPid: () => 456,
      isProcessAlive: () => false,
      now: () => 999,
    });

    const result = runtime.ensureBackgroundStatsServerStarted();

    assert.deepEqual(result, {
      url: 'http://127.0.0.1:6970',
      runningInCurrentProcess: true,
    });
    assert.equal(startedPort, 6970);
    assert.deepEqual(JSON.parse(fs.readFileSync(statePath, 'utf8')), {
      pid: 456,
      port: 6970,
      startedAtMs: 999,
    });
  });
});

test('stats runtime stops owned server and clears daemon state during quit cleanup', async () => {
  await withTempDir(async (dir) => {
    const statePath = path.join(dir, 'stats-daemon.json');
    const calls: string[] = [];

    const runtime = createStatsRuntime({
      statsDaemonStatePath: statePath,
      getResolvedConfig: () => ({
        immersionTracking: { enabled: true },
        stats: { serverPort: 6971 },
      }),
      getImmersionTracker: () => ({}) as never,
      ensureImmersionTrackerStartedCore: () => {},
      startStatsServer: () => ({
        close: () => {
          calls.push('close');
        },
      }),
      openExternal: async () => {},
      exitAppWithCode: () => {},
      destroyStatsWindow: () => {
        calls.push('destroy-window');
      },
      logInfo: () => {},
      logWarn: () => {},
      logError: () => {},
      getCurrentPid: () => 789,
      isProcessAlive: () => true,
      now: () => 500,
    });

    runtime.ensureBackgroundStatsServerStarted();
    runtime.cleanupBeforeQuit();

    assert.deepEqual(calls, ['destroy-window', 'close']);
    assert.equal(fs.existsSync(statePath), false);
    assert.equal(runtime.getStatsServer(), null);
  });
});

test('stats runtime stops the in-process background server without signalling the current process', async () => {
  await withTempDir(async (dir) => {
    const statePath = path.join(dir, 'stats-daemon.json');
    const calls: string[] = [];

    const runtime = createStatsRuntime({
      statsDaemonStatePath: statePath,
      getResolvedConfig: () => ({
        immersionTracking: { enabled: true },
        stats: { serverPort: 6972 },
      }),
      getImmersionTracker: () => ({}) as never,
      ensureImmersionTrackerStartedCore: () => {},
      startStatsServer: () => ({
        close: () => {
          calls.push('close');
        },
      }),
      openExternal: async () => {},
      exitAppWithCode: () => {},
      logInfo: () => {},
      logWarn: () => {},
      logError: () => {},
      getCurrentPid: () => 321,
      isProcessAlive: () => true,
      killProcess: () => {
        calls.push('kill');
      },
      now: () => 500,
    });

    runtime.ensureBackgroundStatsServerStarted();
    const result = await runtime.stopBackgroundStatsServer();

    assert.deepEqual(result, { ok: true, stale: false });
    assert.deepEqual(calls, ['close']);
    assert.equal(fs.existsSync(statePath), false);
    assert.equal(runtime.getStatsServer(), null);
  });
});
