import test from 'node:test';
import assert from 'node:assert/strict';
import { createRunStatsCliCommandHandler } from './stats-cli-command';

function makeHandler(overrides: Partial<Parameters<typeof createRunStatsCliCommandHandler>[0]> = {}) {
  const calls: string[] = [];
  const responses: Array<{ responsePath: string; payload: { ok: boolean; url?: string; error?: string } }> = [];

  const handler = createRunStatsCliCommandHandler({
    getResolvedConfig: () => ({
      immersionTracking: { enabled: true },
      stats: { serverPort: 6969 },
    }),
    ensureImmersionTrackerStarted: () => {
      calls.push('ensureImmersionTrackerStarted');
    },
    getImmersionTracker: () => ({ cleanupVocabularyStats: undefined }),
    ensureStatsServerStarted: () => {
      calls.push('ensureStatsServerStarted');
      return 'http://127.0.0.1:6969';
    },
    openExternal: async (url) => {
      calls.push(`openExternal:${url}`);
    },
    writeResponse: (responsePath, payload) => {
      responses.push({ responsePath, payload });
    },
    exitAppWithCode: (code) => {
      calls.push(`exitAppWithCode:${code}`);
    },
    logInfo: (message) => {
      calls.push(`info:${message}`);
    },
    logWarn: (message) => {
      calls.push(`warn:${message}`);
    },
    logError: (message, error) => {
      calls.push(`error:${message}:${error instanceof Error ? error.message : String(error)}`);
    },
    ...overrides,
  });

  return { handler, calls, responses };
}

test('stats cli command starts tracker, server, browser, and writes success response', async () => {
  const { handler, calls, responses } = makeHandler();

  await handler({ statsResponsePath: '/tmp/subminer-stats-response.json' }, 'initial');

  assert.deepEqual(calls, [
    'ensureImmersionTrackerStarted',
    'ensureStatsServerStarted',
    'openExternal:http://127.0.0.1:6969',
    'info:Stats dashboard available at http://127.0.0.1:6969',
  ]);
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: true, url: 'http://127.0.0.1:6969' },
    },
  ]);
});

test('stats cli command fails when immersion tracking is disabled', async () => {
  const { handler, calls, responses } = makeHandler({
    getResolvedConfig: () => ({
      immersionTracking: { enabled: false },
      stats: { serverPort: 6969 },
    }),
  });

  await handler({ statsResponsePath: '/tmp/subminer-stats-response.json' }, 'initial');

  assert.equal(calls.includes('ensureImmersionTrackerStarted'), false);
  assert.ok(calls.includes('exitAppWithCode:1'));
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: false, error: 'Immersion tracking is disabled in config.' },
    },
  ]);
});

test('stats cli command runs vocab cleanup instead of opening dashboard when cleanup mode is requested', async () => {
  const { handler, calls, responses } = makeHandler({
    getImmersionTracker: () => ({
      cleanupVocabularyStats: async () => ({ scanned: 3, kept: 1, deleted: 2, repaired: 1 }),
    }),
  });

  await handler(
    {
      statsResponsePath: '/tmp/subminer-stats-response.json',
      statsCleanup: true,
      statsCleanupVocab: true,
    },
    'initial',
  );

  assert.deepEqual(calls, [
    'ensureImmersionTrackerStarted',
    'info:Stats vocabulary cleanup complete: scanned=3 kept=1 deleted=2 repaired=1',
  ]);
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: true },
    },
  ]);
});
