import assert from 'node:assert/strict';
import test from 'node:test';
import { composeStatsStartupRuntime } from './stats-startup-composer';

test('composeStatsStartupRuntime returns stats startup handlers', async () => {
  const composed = composeStatsStartupRuntime({
    ensureStatsServerStarted: () => 'http://127.0.0.1:8766',
    ensureBackgroundStatsServerStarted: () => ({
      url: 'http://127.0.0.1:8766',
      runningInCurrentProcess: true,
    }),
    stopBackgroundStatsServer: async () => ({ ok: true, stale: false }),
    ensureImmersionTrackerStarted: () => {},
  });

  assert.equal(composed.ensureStatsServerStarted(), 'http://127.0.0.1:8766');
  assert.deepEqual(composed.ensureBackgroundStatsServerStarted(), {
    url: 'http://127.0.0.1:8766',
    runningInCurrentProcess: true,
  });
  assert.deepEqual(await composed.stopBackgroundStatsServer(), { ok: true, stale: false });
  assert.equal(typeof composed.ensureImmersionTrackerStarted, 'function');
});
