import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
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

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-stats-runtime-test-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
}

async function waitForPendingAnimeMetadata(
  tracker: import('../../core/services/immersion-tracker-service').ImmersionTrackerService,
): Promise<void> {
  const privateApi = tracker as unknown as {
    sessionState: { videoId: number } | null;
    pendingAnimeMetadataUpdates?: Map<number, Promise<void>>;
  };
  const videoId = privateApi.sessionState?.videoId;
  if (!videoId) return;
  await privateApi.pendingAnimeMetadataUpdates?.get(videoId);
}

test('tracker rebuildLifetimeSummaries backfills retained sessions and is idempotent', async () => {
  const dbPath = makeDbPath();
  const previousNowMs = globalThis.__subminerTestNowMs;
  let tracker:
    | import('../../core/services/immersion-tracker-service').ImmersionTrackerService
    | null = null;
  let tracker2:
    | import('../../core/services/immersion-tracker-service').ImmersionTrackerService
    | null = null;
  let tracker3:
    | import('../../core/services/immersion-tracker-service').ImmersionTrackerService
    | null = null;
  const { ImmersionTrackerService } = await import('../../core/services/immersion-tracker-service');
  const { Database } = await import('../../core/services/immersion-tracker/sqlite');

  try {
    globalThis.__subminerTestNowMs = 1_700_000_000;
    tracker = new ImmersionTrackerService({ dbPath });
    tracker.handleMediaChange('/tmp/Frieren S01E01.mkv', 'Episode 1');
    await waitForPendingAnimeMetadata(tracker);
    tracker.recordCardsMined(2);
    tracker.recordSubtitleLine('first line', 0, 1);
    globalThis.__subminerTestNowMs = 1_700_001_000;
    tracker.destroy();
    tracker = null;

    globalThis.__subminerTestNowMs = 1_700_002_000;
    tracker2 = new ImmersionTrackerService({ dbPath });
    tracker2.handleMediaChange('/tmp/Frieren S01E02.mkv', 'Episode 2');
    await waitForPendingAnimeMetadata(tracker2);
    tracker2.recordCardsMined(1);
    tracker2.recordSubtitleLine('second line', 0, 1);
    globalThis.__subminerTestNowMs = 1_700_003_000;
    tracker2.destroy();
    tracker2 = null;

    const beforeDb = new Database(dbPath);
    const expectedGlobal = beforeDb
      .prepare(
        `
      SELECT total_sessions, total_cards, episodes_started, active_days
      FROM imm_lifetime_global
      `,
      )
      .get() as {
      total_sessions: number;
      total_cards: number;
      episodes_started: number;
      active_days: number;
    } | null;
    const expectedAnimeRows = (
      beforeDb.prepare('SELECT COUNT(*) AS total FROM imm_lifetime_anime').get() as {
        total: number;
      }
    ).total;
    const expectedMediaRows = (
      beforeDb.prepare('SELECT COUNT(*) AS total FROM imm_lifetime_media').get() as {
        total: number;
      }
    ).total;
    const expectedAppliedSessions = (
      beforeDb.prepare('SELECT COUNT(*) AS total FROM imm_lifetime_applied_sessions').get() as {
        total: number;
      }
    ).total;

    beforeDb.exec(`
      DELETE FROM imm_lifetime_anime;
      DELETE FROM imm_lifetime_media;
      DELETE FROM imm_lifetime_applied_sessions;
      UPDATE imm_lifetime_global
      SET total_sessions = 999,
          total_cards = 999,
          episodes_started = 999,
          active_days = 999
      WHERE global_id = 1;
    `);
    beforeDb.close();

    globalThis.__subminerTestNowMs = 1_700_004_000;
    tracker3 = new ImmersionTrackerService({ dbPath });
    const firstRebuild = await tracker3.rebuildLifetimeSummaries();
    globalThis.__subminerTestNowMs = 1_700_005_000;
    const secondRebuild = await tracker3.rebuildLifetimeSummaries();

    const rebuiltDb = new Database(dbPath);
    const rebuiltGlobal = rebuiltDb
      .prepare(
        `
      SELECT total_sessions, total_cards, episodes_started, active_days
      FROM imm_lifetime_global
      `,
      )
      .get() as {
      total_sessions: number;
      total_cards: number;
      episodes_started: number;
      active_days: number;
    } | null;
    const rebuiltAnimeRows = (
      rebuiltDb.prepare('SELECT COUNT(*) AS total FROM imm_lifetime_anime').get() as {
        total: number;
      }
    ).total;
    const rebuiltMediaRows = (
      rebuiltDb.prepare('SELECT COUNT(*) AS total FROM imm_lifetime_media').get() as {
        total: number;
      }
    ).total;
    const rebuiltAppliedSessions = (
      rebuiltDb.prepare('SELECT COUNT(*) AS total FROM imm_lifetime_applied_sessions').get() as {
        total: number;
      }
    ).total;
    rebuiltDb.close();

    assert.ok(rebuiltGlobal);
    assert.ok(expectedGlobal);
    assert.equal(rebuiltGlobal?.total_sessions, expectedGlobal?.total_sessions);
    assert.equal(rebuiltGlobal?.total_cards, expectedGlobal?.total_cards);
    assert.equal(rebuiltGlobal?.episodes_started, expectedGlobal?.episodes_started);
    assert.equal(rebuiltGlobal?.active_days, expectedGlobal?.active_days);
    assert.equal(rebuiltAnimeRows, expectedAnimeRows);
    assert.equal(rebuiltMediaRows, expectedMediaRows);
    assert.equal(rebuiltAppliedSessions, expectedAppliedSessions);
    assert.equal(firstRebuild.appliedSessions, expectedAppliedSessions);
    assert.equal(secondRebuild.appliedSessions, firstRebuild.appliedSessions);
    assert.ok(secondRebuild.rebuiltAtMs >= firstRebuild.rebuiltAtMs);
  } finally {
    globalThis.__subminerTestNowMs = previousNowMs;
    tracker?.destroy();
    tracker2?.destroy();
    tracker3?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('stats cli command runs lifetime rebuild when requested', async () => {
  const { handler, calls, responses } = makeHandler({
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
