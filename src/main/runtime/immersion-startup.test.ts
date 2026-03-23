import test from 'node:test';
import assert from 'node:assert/strict';
import { createImmersionTrackerStartupHandler } from './immersion-startup';

function makeConfig() {
  return {
    immersionTracking: {
      enabled: true,
      batchSize: 40,
      flushIntervalMs: 1500,
      queueCap: 500,
      payloadCapBytes: 16000,
      maintenanceIntervalMs: 3600000,
      retention: {
        eventsDays: 14,
        telemetryDays: 30,
        sessionsDays: 45,
        dailyRollupsDays: 180,
        monthlyRollupsDays: 730,
        vacuumIntervalDays: 7,
      },
    },
  };
}

test('createImmersionTrackerStartupHandler skips when disabled', () => {
  const calls: string[] = [];
  let tracker: unknown = 'unchanged';
  const handler = createImmersionTrackerStartupHandler({
    getResolvedConfig: () => ({
      immersionTracking: {
        ...makeConfig().immersionTracking,
        enabled: false,
      },
    }),
    getConfiguredDbPath: () => '/tmp/subminer.db',
    createTrackerService: () => {
      calls.push('createTrackerService');
      return {};
    },
    setTracker: (nextTracker) => {
      tracker = nextTracker;
    },
    getMpvClient: () => null,
    seedTrackerFromCurrentMedia: () => calls.push('seedTracker'),
    logInfo: (message) => calls.push(`info:${message}`),
    logDebug: (message) => calls.push(`debug:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
  });

  handler();

  assert.ok(calls.includes('info:Immersion tracking disabled in config'));
  assert.equal(calls.includes('createTrackerService'), false);
  assert.equal(calls.includes('seedTracker'), false);
  assert.equal(tracker, 'unchanged');
});

test('createImmersionTrackerStartupHandler skips when env disables session tracking', () => {
  const calls: string[] = [];
  const originalEnv = process.env.SUBMINER_DISABLE_IMMERSION_TRACKING;
  process.env.SUBMINER_DISABLE_IMMERSION_TRACKING = '1';

  try {
    let tracker: unknown = 'unchanged';
    const handler = createImmersionTrackerStartupHandler({
      getResolvedConfig: () => {
        calls.push('getResolvedConfig');
        return makeConfig();
      },
      getConfiguredDbPath: () => {
        calls.push('getConfiguredDbPath');
        return '/tmp/subminer.db';
      },
      createTrackerService: () => {
        calls.push('createTrackerService');
        return {};
      },
      setTracker: (nextTracker) => {
        tracker = nextTracker;
      },
      getMpvClient: () => null,
      seedTrackerFromCurrentMedia: () => calls.push('seedTracker'),
      logInfo: (message) => calls.push(`info:${message}`),
      logDebug: (message) => calls.push(`debug:${message}`),
      logWarn: (message) => calls.push(`warn:${message}`),
    });

    handler();

    assert.equal(calls.includes('getResolvedConfig'), false);
    assert.equal(calls.includes('getConfiguredDbPath'), false);
    assert.equal(calls.includes('createTrackerService'), false);
    assert.equal(calls.includes('seedTracker'), false);
    assert.equal(tracker, 'unchanged');
    assert.ok(
      calls.includes(
        'info:Immersion tracking disabled for this session by SUBMINER_DISABLE_IMMERSION_TRACKING=1.',
      ),
    );
  } finally {
    if (originalEnv === undefined) {
      delete process.env.SUBMINER_DISABLE_IMMERSION_TRACKING;
    } else {
      process.env.SUBMINER_DISABLE_IMMERSION_TRACKING = originalEnv;
    }
  }
});

test('createImmersionTrackerStartupHandler creates tracker and auto-connects mpv', () => {
  const calls: string[] = [];
  const trackerInstance = { kind: 'tracker' };
  let assignedTracker: unknown = null;
  let receivedDbPath = '';
  let receivedPolicy: unknown;
  let connectCalls = 0;
  const handler = createImmersionTrackerStartupHandler({
    getResolvedConfig: () => makeConfig(),
    getConfiguredDbPath: () => '/tmp/subminer.db',
    createTrackerService: (params) => {
      receivedDbPath = params.dbPath;
      receivedPolicy = params.policy;
      return trackerInstance;
    },
    setTracker: (nextTracker) => {
      assignedTracker = nextTracker;
    },
    getMpvClient: () => ({
      connected: false,
      connect: () => {
        connectCalls += 1;
      },
    }),
    seedTrackerFromCurrentMedia: () => calls.push('seedTracker'),
    logInfo: (message) => calls.push(`info:${message}`),
    logDebug: (message) => calls.push(`debug:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
  });

  handler();

  assert.equal(receivedDbPath, '/tmp/subminer.db');
  assert.deepEqual(receivedPolicy, {
    batchSize: 40,
    flushIntervalMs: 1500,
    queueCap: 500,
    payloadCapBytes: 16000,
    maintenanceIntervalMs: 3600000,
    retention: {
      eventsDays: 14,
      telemetryDays: 30,
      sessionsDays: 45,
      dailyRollupsDays: 180,
      monthlyRollupsDays: 730,
      vacuumIntervalDays: 7,
    },
  });
  assert.equal(assignedTracker, trackerInstance);
  assert.equal(connectCalls, 1);
  assert.ok(calls.includes('seedTracker'));
  assert.ok(calls.includes('info:Auto-connecting MPV client for immersion tracking'));
});

test('createImmersionTrackerStartupHandler disables tracker on failure', () => {
  const calls: string[] = [];
  let assignedTracker: unknown = 'initial';
  const handler = createImmersionTrackerStartupHandler({
    getResolvedConfig: () => makeConfig(),
    getConfiguredDbPath: () => '/tmp/subminer.db',
    createTrackerService: () => {
      throw new Error('db unavailable');
    },
    setTracker: (nextTracker) => {
      assignedTracker = nextTracker;
    },
    getMpvClient: () => null,
    seedTrackerFromCurrentMedia: () => calls.push('seedTracker'),
    logInfo: (message) => calls.push(`info:${message}`),
    logDebug: (message) => calls.push(`debug:${message}`),
    logWarn: (message, details) => calls.push(`warn:${message}:${(details as Error).message}`),
  });

  handler();

  assert.equal(assignedTracker, null);
  assert.equal(calls.includes('seedTracker'), false);
  assert.ok(
    calls.includes('warn:Immersion tracker startup failed; disabling tracking.:db unavailable'),
  );
});

test('createImmersionTrackerStartupHandler skips mpv auto-connect when disabled by caller', () => {
  let connectCalls = 0;
  const handler = createImmersionTrackerStartupHandler({
    getResolvedConfig: () => makeConfig(),
    getConfiguredDbPath: () => '/tmp/subminer.db',
    createTrackerService: () => ({}),
    setTracker: () => {},
    getMpvClient: () => ({
      connected: false,
      connect: () => {
        connectCalls += 1;
      },
    }),
    shouldAutoConnectMpv: () => false,
    seedTrackerFromCurrentMedia: () => {},
    logInfo: () => {},
    logDebug: () => {},
    logWarn: () => {},
  });

  handler();

  assert.equal(connectCalls, 0);
});
