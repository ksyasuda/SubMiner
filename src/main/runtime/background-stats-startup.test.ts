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

test('ensures background stats server and logs local startup', () => {
  const { deps, calls } = createDeps();

  createEnsureBackgroundStatsServerHandler(deps)();

  assert.ok(calls.includes('ensureBackgroundStatsServerStarted'));
  assert.ok(
    calls.some((value) => value.startsWith('info:') && value.includes('http://127.0.0.1:3888')),
  );
});

test('logs reuse when a background stats server is already running', () => {
  const { deps, calls } = createDeps({
    ensureBackgroundStatsServerStarted: () => ({
      url: 'http://127.0.0.1:3888',
      runningInCurrentProcess: false,
    }),
  });

  createEnsureBackgroundStatsServerHandler(deps)();

  assert.ok(
    calls.some((value) => value.startsWith('info:') && /already running|reusing/i.test(value)),
  );
});

test('skips when stats.autoStartServer is disabled', () => {
  const { deps, calls } = createDeps({ isStatsAutoStartEnabled: () => false });

  createEnsureBackgroundStatsServerHandler(deps)();

  assert.equal(calls.includes('ensureBackgroundStatsServerStarted'), false);
});

test('skips when immersion tracking is disabled', () => {
  const { deps, calls } = createDeps({ isImmersionTrackingEnabled: () => false });

  createEnsureBackgroundStatsServerHandler(deps)();

  assert.equal(calls.includes('ensureBackgroundStatsServerStarted'), false);
});

test('logs a warning instead of throwing when startup fails', () => {
  const { deps, calls } = createDeps({
    ensureBackgroundStatsServerStarted: () => {
      throw new Error('port in use');
    },
  });

  assert.doesNotThrow(() => createEnsureBackgroundStatsServerHandler(deps)());
  assert.ok(calls.some((value) => value.startsWith('warn:')));
});
