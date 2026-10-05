import test from 'node:test';
import assert from 'node:assert/strict';
import { createRunStatsCliCommandHandler } from './stats-cli-command';

function makeHandler(
  overrides: Partial<Parameters<typeof createRunStatsCliCommandHandler>[0]> = {},
) {
  const calls: string[] = [];
  const responses: Array<{
    responsePath: string;
    payload: { ok: boolean; url?: string; error?: string };
  }> = [];

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
    ensureBackgroundStatsServerStarted: () => ({
      url: 'http://127.0.0.1:6969',
      runningInCurrentProcess: true,
    }),
    stopBackgroundStatsServer: async () => ({ ok: true, stale: false }),
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

test('stats cli command respects stats.autoOpenBrowser=false', async () => {
  const { handler, calls, responses } = makeHandler({
    getResolvedConfig: () => ({
      immersionTracking: { enabled: true },
      stats: { serverPort: 6969, autoOpenBrowser: false },
    }),
  });

  await handler({ statsResponsePath: '/tmp/subminer-stats-response.json' }, 'initial');

  assert.deepEqual(calls, [
    'ensureImmersionTrackerStarted',
    'ensureStatsServerStarted',
    'info:Stats dashboard available at http://127.0.0.1:6969',
  ]);
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: true, url: 'http://127.0.0.1:6969' },
    },
  ]);
});

test('stats cli command starts background daemon without opening browser', async () => {
  const { handler, calls, responses } = makeHandler({
    ensureBackgroundStatsServerStarted: () => {
      calls.push('ensureBackgroundStatsServerStarted');
      return { url: 'http://127.0.0.1:6969', runningInCurrentProcess: true };
    },
  } as never);

  await handler(
    {
      statsResponsePath: '/tmp/subminer-stats-response.json',
      statsBackground: true,
    } as never,
    'initial',
  );

  assert.deepEqual(calls, [
    'ensureBackgroundStatsServerStarted',
    'info:Stats dashboard available at http://127.0.0.1:6969',
  ]);
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: true, url: 'http://127.0.0.1:6969' },
    },
  ]);
});

test('stats cli command exits helper app when background daemon is already running elsewhere', async () => {
  const { handler, calls, responses } = makeHandler({
    ensureBackgroundStatsServerStarted: () => {
      calls.push('ensureBackgroundStatsServerStarted');
      return { url: 'http://127.0.0.1:6969', runningInCurrentProcess: false };
    },
  } as never);

  await handler(
    {
      statsResponsePath: '/tmp/subminer-stats-response.json',
      statsBackground: true,
    } as never,
    'initial',
  );

  assert.ok(calls.includes('exitAppWithCode:0'));
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: true, url: 'http://127.0.0.1:6969' },
    },
  ]);
});

test('stats cli command stops background daemon and treats stale state as success', async () => {
  const { handler, calls, responses } = makeHandler({
    stopBackgroundStatsServer: async () => {
      calls.push('stopBackgroundStatsServer');
      return { ok: true, stale: true };
    },
  } as never);

  await handler(
    {
      statsResponsePath: '/tmp/subminer-stats-response.json',
      statsStop: true,
    } as never,
    'initial',
  );

  assert.deepEqual(calls, [
    'stopBackgroundStatsServer',
    'info:Background stats server is not running; cleaned stale state.',
    'exitAppWithCode:0',
  ]);
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: true },
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

test('stats cli command runs a duplicate-line cleanup preview without touching the dashboard', async () => {
  const { handler, calls, responses } = makeHandler({
    getImmersionTracker: () => ({
      cleanupDuplicateSubtitleLines: async (options: {
        dryRun?: boolean;
        lookbackDays?: number | null;
      }) => ({
        dryRun: options.dryRun === true,
        lookbackDays: options.lookbackDays ?? null,
        scannedLines: 900,
        burstGroups: 2,
        removedLines: 180,
        removedWordOccurrences: 540,
        removedKanjiOccurrences: 120,
        samples: [
          {
            videoId: 7,
            videoTitle: 'Ep 1',
            text: '飛び上がる',
            frames: 90,
            removedLines: 89,
            startMs: 1000,
            endMs: 5000,
          },
        ],
      }),
    }),
  });

  await handler(
    {
      statsResponsePath: '/tmp/subminer-stats-response.json',
      statsCleanup: true,
      statsCleanupDuplicateLines: true,
      statsCleanupDryRun: true,
      statsCleanupLookbackDays: 30,
    },
    'initial',
  );

  assert.deepEqual(calls, [
    'ensureImmersionTrackerStarted',
    'info:Stats duplicate-line cleanup preview (last 30d): scanned=900 bursts=2 removedLines=180 removedWordCounts=540 removedKanjiCounts=120',
    'info:  Ep 1: "飛び上がる" x90',
  ]);
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: true },
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

test('stats cli command runs lifetime rebuild when cleanup lifetime mode is requested', async () => {
  const { handler, calls, responses } = makeHandler({
    ensureVocabularyCleanupTokenizerReady: async () => {
      calls.push('ensureVocabularyCleanupTokenizerReady');
    },
    getImmersionTracker: () => ({
      rebuildLifetimeSummaries: async () => ({
        appliedSessions: 4,
        rebuiltAtMs: 1_710_000_000,
      }),
    }),
  });

  await handler(
    {
      statsResponsePath: '/tmp/subminer-stats-response.json',
      statsCleanup: true,
      statsCleanupLifetime: true,
    },
    'initial',
  );

  assert.deepEqual(calls, [
    'ensureImmersionTrackerStarted',
    'info:Stats lifetime rebuild complete: appliedSessions=4 rebuiltAtMs=1710000000',
  ]);
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: true },
    },
  ]);
});

test('stats cli command rejects cleanup calls without exactly one cleanup mode', async () => {
  const { handler, calls, responses } = makeHandler({
    getImmersionTracker: () => ({
      cleanupVocabularyStats: async () => ({ scanned: 1, kept: 1, deleted: 0, repaired: 0 }),
      rebuildLifetimeSummaries: async () => ({ appliedSessions: 0, rebuiltAtMs: 0 }),
    }),
  });

  await handler(
    {
      statsResponsePath: '/tmp/subminer-stats-response.json',
      statsCleanup: true,
      statsCleanupVocab: true,
      statsCleanupLifetime: true,
    },
    'initial',
  );

  assert.ok(calls.includes('error:Stats command failed:Choose exactly one stats cleanup mode.'));
  assert.deepEqual(responses, [
    {
      responsePath: '/tmp/subminer-stats-response.json',
      payload: { ok: false, error: 'Choose exactly one stats cleanup mode.' },
    },
  ]);
});
