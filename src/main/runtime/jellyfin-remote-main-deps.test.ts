import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler,
  createBuildHandleJellyfinRemotePlayMainDepsHandler,
  createBuildHandleJellyfinRemotePlaystateMainDepsHandler,
  createBuildReportJellyfinRemoteProgressMainDepsHandler,
  createBuildReportJellyfinRemoteStoppedMainDepsHandler,
} from './jellyfin-remote-main-deps';

test('jellyfin remote play main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildHandleJellyfinRemotePlayMainDepsHandler({
    getConfiguredSession: () => ({ id: 1 }) as never,
    getClientInfo: () => ({ id: 2 }) as never,
    getJellyfinConfig: () => ({ id: 3 }),
    playJellyfinItem: async () => {
      calls.push('play');
    },
    logWarn: (message) => calls.push(`warn:${message}`),
  })();

  assert.deepEqual(deps.getConfiguredSession(), { id: 1 });
  assert.deepEqual(deps.getClientInfo(), { id: 2 });
  assert.deepEqual(deps.getJellyfinConfig(), { id: 3 });
  await deps.playJellyfinItem({} as never);
  deps.logWarn('missing');
  assert.deepEqual(calls, ['play', 'warn:missing']);
});

test('jellyfin remote playstate main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildHandleJellyfinRemotePlaystateMainDepsHandler({
    getMpvClient: () => ({ id: 1 }),
    sendMpvCommand: () => calls.push('send'),
    reportJellyfinRemoteProgress: async () => {
      calls.push('progress');
    },
    reportJellyfinRemoteStopped: async () => {
      calls.push('stopped');
    },
    jellyfinTicksToSeconds: (ticks) => ticks / 10,
  })();

  assert.deepEqual(deps.getMpvClient(), { id: 1 });
  deps.sendMpvCommand({} as never, ['stop']);
  await deps.reportJellyfinRemoteProgress(true);
  await deps.reportJellyfinRemoteStopped();
  assert.equal(deps.jellyfinTicksToSeconds(100), 10);
  assert.deepEqual(calls, ['send', 'progress', 'stopped']);
});

test('jellyfin remote general command main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const playback = { itemId: 'abc', playMethod: 'DirectPlay' as const };
  const deps = createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler({
    getMpvClient: () => ({ id: 1 }),
    sendMpvCommand: () => calls.push('send'),
    getActivePlayback: () => playback,
    reportJellyfinRemoteProgress: async () => {
      calls.push('progress');
    },
    logDebug: (message) => calls.push(`debug:${message}`),
  })();

  assert.deepEqual(deps.getMpvClient(), { id: 1 });
  deps.sendMpvCommand({} as never, ['set_property', 'sid', 1]);
  assert.deepEqual(deps.getActivePlayback(), playback);
  await deps.reportJellyfinRemoteProgress(true);
  deps.logDebug('ignore');
  assert.deepEqual(calls, ['send', 'progress', 'debug:ignore']);
});

test('jellyfin remote progress main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildReportJellyfinRemoteProgressMainDepsHandler({
    getActivePlayback: () => ({ itemId: 'abc', playMethod: 'DirectPlay' }),
    clearActivePlayback: () => calls.push('clear'),
    getSession: () => ({ id: 1, isConnected: () => true }) as never,
    getMpvClient: () => ({ id: 2, requestProperty: async () => 0 }) as never,
    getNow: () => 123,
    getLastProgressAtMs: () => 10,
    setLastProgressAtMs: () => calls.push('set-last'),
    progressIntervalMs: 2500,
    ticksPerSecond: 10000000,
    logDebug: (message) => calls.push(`debug:${message}`),
  })();

  assert.equal(deps.getNow(), 123);
  assert.equal(deps.getLastProgressAtMs(), 10);
  deps.setLastProgressAtMs(5);
  assert.equal(deps.progressIntervalMs, 2500);
  assert.equal(deps.ticksPerSecond, 10000000);
  deps.clearActivePlayback();
  deps.logDebug('x', null);
  assert.deepEqual(calls, ['set-last', 'clear', 'debug:x']);
});

test('jellyfin remote stopped main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const session = { id: 1, isConnected: () => true };
  const deps = createBuildReportJellyfinRemoteStoppedMainDepsHandler({
    getActivePlayback: () => ({ itemId: 'abc', playMethod: 'DirectPlay' }),
    clearActivePlayback: () => calls.push('clear'),
    getSession: () => session as never,
    getMpvClient: () => ({ id: 2, currentTimePos: 4 }) as never,
    ticksPerSecond: 10_000_000,
    logDebug: (message) => calls.push(`debug:${message}`),
  })();

  assert.deepEqual(deps.getActivePlayback(), { itemId: 'abc', playMethod: 'DirectPlay' });
  deps.clearActivePlayback();
  assert.equal(deps.getSession(), session);
  assert.deepEqual(deps.getMpvClient(), { id: 2, currentTimePos: 4 });
  assert.equal(deps.ticksPerSecond, 10_000_000);
  deps.logDebug('stopped', null);
  assert.deepEqual(calls, ['clear', 'debug:stopped']);
});
