import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildAnilistStateRuntimeMainDepsHandler,
  createBuildConfigDerivedRuntimeMainDepsHandler,
  createBuildImmersionMediaRuntimeMainDepsHandler,
  createBuildMainSubsyncRuntimeMainDepsHandler,
} from './runtime-bootstrap-main-deps';

test('immersion media runtime main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildImmersionMediaRuntimeMainDepsHandler({
    getResolvedConfig: () => ({ immersionTracking: { dbPath: '/tmp/db.sqlite' } }),
    defaultImmersionDbPath: '/tmp/default.sqlite',
    getTracker: () => ({ handleMediaChange: () => calls.push('track') }),
    getMpvClient: () => ({ connected: true }),
    getCurrentMediaPath: () => '/tmp/media.mkv',
    getCurrentMediaTitle: () => 'Title',
    sleep: async () => {
      calls.push('sleep');
    },
    seedWaitMs: 25,
    seedAttempts: 3,
    logDebug: (message) => calls.push(`debug:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
  })();

  assert.equal(deps.defaultImmersionDbPath, '/tmp/default.sqlite');
  assert.deepEqual(deps.getResolvedConfig(), { immersionTracking: { dbPath: '/tmp/db.sqlite' } });
  assert.deepEqual(deps.getMpvClient(), { connected: true });
  assert.equal(deps.getCurrentMediaPath(), '/tmp/media.mkv');
  assert.equal(deps.getCurrentMediaTitle(), 'Title');
  assert.equal(deps.seedWaitMs, 25);
  assert.equal(deps.seedAttempts, 3);
  await deps.sleep?.(1);
  deps.logDebug('a');
  deps.logInfo('b');
  assert.deepEqual(calls, ['sleep', 'debug:a', 'info:b']);
});

test('anilist state runtime main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildAnilistStateRuntimeMainDepsHandler({
    getClientSecretState: () => ({ status: 'resolved' } as never),
    setClientSecretState: () => calls.push('set-client'),
    getRetryQueueState: () => ({ pending: 1 } as never),
    setRetryQueueState: () => calls.push('set-queue'),
    getUpdateQueueSnapshot: () => ({ pending: 2 } as never),
    clearStoredToken: () => calls.push('clear-stored'),
    clearCachedAccessToken: () => calls.push('clear-cached'),
  })();

  assert.deepEqual(deps.getClientSecretState(), { status: 'resolved' });
  assert.deepEqual(deps.getRetryQueueState(), { pending: 1 });
  assert.deepEqual(deps.getUpdateQueueSnapshot(), { pending: 2 });
  deps.setClientSecretState({} as never);
  deps.setRetryQueueState({} as never);
  deps.clearStoredToken();
  deps.clearCachedAccessToken();
  assert.deepEqual(calls, ['set-client', 'set-queue', 'clear-stored', 'clear-cached']);
});

test('config derived runtime main deps builder maps callbacks', () => {
  const deps = createBuildConfigDerivedRuntimeMainDepsHandler({
    getResolvedConfig: () => ({ jimaku: {} } as never),
    getRuntimeOptionsManager: () => null,
    platform: 'darwin',
    defaultJimakuLanguagePreference: 'ja',
    defaultJimakuMaxEntryResults: 20,
    defaultJimakuApiBaseUrl: 'https://api.example.com',
  })();

  assert.deepEqual(deps.getResolvedConfig(), { jimaku: {} });
  assert.equal(deps.getRuntimeOptionsManager(), null);
  assert.equal(deps.platform, 'darwin');
  assert.equal(deps.defaultJimakuLanguagePreference, 'ja');
  assert.equal(deps.defaultJimakuMaxEntryResults, 20);
  assert.equal(deps.defaultJimakuApiBaseUrl, 'https://api.example.com');
});

test('main subsync runtime main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildMainSubsyncRuntimeMainDepsHandler({
    getMpvClient: () => ({ connected: true }) as never,
    getResolvedConfig: () => ({ subsync: {} } as never),
    getSubsyncInProgress: () => true,
    setSubsyncInProgress: () => calls.push('set-progress'),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    openManualPicker: () => calls.push('open-picker'),
  })();

  assert.deepEqual(deps.getMpvClient(), { connected: true });
  assert.deepEqual(deps.getResolvedConfig(), { subsync: {} });
  assert.equal(deps.getSubsyncInProgress(), true);
  deps.setSubsyncInProgress(false);
  deps.showMpvOsd('ready');
  deps.openManualPicker({} as never);
  assert.deepEqual(calls, ['set-progress', 'osd:ready', 'open-picker']);
});
