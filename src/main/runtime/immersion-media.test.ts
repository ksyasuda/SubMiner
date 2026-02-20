import test from 'node:test';
import assert from 'node:assert/strict';
import { createImmersionMediaRuntime } from './immersion-media';

test('getConfiguredDbPath uses trimmed configured path with fallback', () => {
  const runtime = createImmersionMediaRuntime({
    getResolvedConfig: () => ({ immersionTracking: { dbPath: '  /tmp/custom.db  ' } }),
    defaultImmersionDbPath: '/tmp/default.db',
    getTracker: () => null,
    getMpvClient: () => null,
    getCurrentMediaPath: () => null,
    getCurrentMediaTitle: () => null,
    logDebug: () => {},
    logInfo: () => {},
  });
  assert.equal(runtime.getConfiguredDbPath(), '/tmp/custom.db');

  const fallbackRuntime = createImmersionMediaRuntime({
    getResolvedConfig: () => ({ immersionTracking: { dbPath: '   ' } }),
    defaultImmersionDbPath: '/tmp/default.db',
    getTracker: () => null,
    getMpvClient: () => null,
    getCurrentMediaPath: () => null,
    getCurrentMediaTitle: () => null,
    logDebug: () => {},
    logInfo: () => {},
  });
  assert.equal(fallbackRuntime.getConfiguredDbPath(), '/tmp/default.db');
});

test('syncFromCurrentMediaState uses current media path directly', () => {
  const calls: Array<{ path: string; title: string | null }> = [];
  const runtime = createImmersionMediaRuntime({
    getResolvedConfig: () => ({}),
    defaultImmersionDbPath: '/tmp/default.db',
    getTracker: () => ({
      handleMediaChange: (path, title) => calls.push({ path, title }),
    }),
    getMpvClient: () => ({ connected: true, currentVideoPath: '/tmp/video.mkv' }),
    getCurrentMediaPath: () => ' /tmp/current.mkv ',
    getCurrentMediaTitle: () => '  Current Title  ',
    logDebug: () => {},
    logInfo: () => {},
  });

  runtime.syncFromCurrentMediaState();
  assert.deepEqual(calls, [{ path: '/tmp/current.mkv', title: 'Current Title' }]);
});

test('seedFromCurrentMedia resolves media path from mpv properties', async () => {
  const calls: Array<{ path: string; title: string | null }> = [];
  const runtime = createImmersionMediaRuntime({
    getResolvedConfig: () => ({}),
    defaultImmersionDbPath: '/tmp/default.db',
    getTracker: () => ({
      handleMediaChange: (path, title) => calls.push({ path, title }),
    }),
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name: string) => {
        if (name === 'path') return '/tmp/from-property.mkv';
        if (name === 'media-title') return 'Property Title';
        return null;
      },
    }),
    getCurrentMediaPath: () => null,
    getCurrentMediaTitle: () => null,
    sleep: async () => {},
    seedAttempts: 2,
    logDebug: () => {},
    logInfo: () => {},
  });

  await runtime.seedFromCurrentMedia();
  assert.deepEqual(calls, [{ path: '/tmp/from-property.mkv', title: 'Property Title' }]);
});
