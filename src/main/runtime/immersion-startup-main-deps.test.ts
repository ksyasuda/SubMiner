import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildImmersionTrackerStartupMainDepsHandler } from './immersion-startup-main-deps';

test('immersion tracker startup main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildImmersionTrackerStartupMainDepsHandler({
    getResolvedConfig: () => ({ immersionTracking: { enabled: true } }) as never,
    getConfiguredDbPath: () => '/tmp/immersion.db',
    createTrackerService: () => {
      calls.push('create');
      return { id: 'tracker' };
    },
    setTracker: () => calls.push('set'),
    getMpvClient: () => ({ connected: true, connect: () => {} }),
    seedTrackerFromCurrentMedia: () => calls.push('seed'),
    logInfo: (message) => calls.push(`info:${message}`),
    logDebug: (message) => calls.push(`debug:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
  })();

  assert.deepEqual(deps.getResolvedConfig(), { immersionTracking: { enabled: true } });
  assert.equal(deps.getConfiguredDbPath(), '/tmp/immersion.db');
  assert.deepEqual(
    deps.createTrackerService({ dbPath: '/tmp/immersion.db', policy: {} as never }),
    {
      id: 'tracker',
    },
  );
  deps.setTracker(null);
  assert.equal(deps.getMpvClient()?.connected, true);
  deps.seedTrackerFromCurrentMedia();
  deps.logInfo('i');
  deps.logDebug('d');
  deps.logWarn('w', null);

  assert.deepEqual(calls, ['create', 'set', 'seed', 'info:i', 'debug:d', 'warn:w']);
});
