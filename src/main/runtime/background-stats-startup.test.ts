import test from 'node:test';
import assert from 'node:assert/strict';
import { createEnsureBackgroundStatsServerHandler } from './background-stats-startup';

function createDeps(
  overrides: Partial<Parameters<typeof createEnsureBackgroundStatsServerHandler>[0]> = {},
) {
  const calls: string[] = [];
  const deps: Parameters<typeof createEnsureBackgroundStatsServerHandler>[0] = {
    isStatsAutoStartEnabled: () => true,
    isImmersionTrackingEnabled: () => true,
    ensureBackgroundStatsServerStarted: () => {
      calls.push('ensureBackgroundStatsServerStarted');
      return { url: 'http://127.0.0.1:3888', runningInCurrentProcess: true };
    },
    logInfo: (message) => {
      calls.push(`info:${message}`);
    },
    logWarn: (message) => {
      calls.push(`warn:${message}`);
    },
    ...overrides,
  };
  return { deps, calls };
}

test('ensures background stats server and logs local startup', async () => {
  const { deps, calls } = createDeps();

  await createEnsureBackgroundStatsServerHandler(deps)();

  assert.ok(calls.includes('ensureBackgroundStatsServerStarted'));
  assert.ok(
    calls.some((value) => value.startsWith('info:') && value.includes('http://127.0.0.1:3888')),
  );
});

test('logs reuse when a background stats server is already running', async () => {
  const { deps, calls } = createDeps({
    ensureBackgroundStatsServerStarted: () => ({
      url: 'http://127.0.0.1:3888',
      runningInCurrentProcess: false,
    }),
  });

  await createEnsureBackgroundStatsServerHandler(deps)();

  assert.ok(
    calls.some((value) => value.startsWith('info:') && /already running|reusing/i.test(value)),
  );
});

test('skips when stats.autoStartServer is disabled', async () => {
  const { deps, calls } = createDeps({ isStatsAutoStartEnabled: () => false });

  await createEnsureBackgroundStatsServerHandler(deps)();

  assert.equal(calls.includes('ensureBackgroundStatsServerStarted'), false);
});

test('skips when immersion tracking is disabled', async () => {
  const { deps, calls } = createDeps({ isImmersionTrackingEnabled: () => false });

  await createEnsureBackgroundStatsServerHandler(deps)();

  assert.equal(calls.includes('ensureBackgroundStatsServerStarted'), false);
});

test('logs a warning instead of throwing when startup fails', async () => {
  const { deps, calls } = createDeps({
    ensureBackgroundStatsServerStarted: () => {
      throw new Error('port in use');
    },
  });

  await assert.doesNotReject(createEnsureBackgroundStatsServerHandler(deps)());
  assert.ok(calls.some((value) => value.startsWith('warn:')));
});

test('logs an asynchronously reported startup failure', async () => {
  const { deps, calls } = createDeps({
    ensureBackgroundStatsServerStarted: async () => {
      await Promise.resolve();
      throw new Error('address in use');
    },
  });

  await createEnsureBackgroundStatsServerHandler(deps)();

  assert.ok(calls.some((value) => value.startsWith('warn:')));
  assert.equal(
    calls.some((value) => value.startsWith('info:')),
    false,
  );
});
