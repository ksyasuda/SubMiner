import test from 'node:test';
import assert from 'node:assert/strict';
import {
  ensureLauncherSetupReady,
  resolveLauncherGateBackend,
  waitForSetupCompletion,
} from './setup-gate';
import { createDefaultSetupState, type SetupState } from '../src/shared/setup-state';

/** An in-progress setup state; tests override only the fields that matter. */
function makeState(overrides: Partial<SetupState> = {}): SetupState {
  return {
    version: 4,
    status: 'in_progress',
    completedAt: null,
    completionSource: null,
    yomitanSetupMode: null,
    lastSeenYomitanDictionaryCount: 0,
    pluginInstallStatus: 'unknown',
    pluginInstallPathSummary: null,
    windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
    windowsMpvShortcutLastStatus: 'unknown',
    bunInstallStatus: 'unknown',
    launcherInstallStatus: 'unknown',
    launcherInstallPath: null,
    ...overrides,
  };
}

function completedState(overrides: Partial<SetupState> = {}): SetupState {
  return makeState({
    status: 'completed',
    completedAt: '2026-03-07T00:00:00.000Z',
    completionSource: 'user',
    yomitanSetupMode: 'internal',
    lastSeenYomitanDictionaryCount: 1,
    ...overrides,
  });
}

const cancelledState = () => makeState({ status: 'cancelled' });

/** Deterministic clock: every `now()` advances 100ms and `sleep` returns immediately. */
function fakeClock() {
  let value = 0;
  return {
    sleep: async () => undefined,
    now: () => (value += 100),
    timeoutMs: 5_000,
    pollIntervalMs: 100,
  };
}

/** Reader that serves `states` in order, then repeats the last one; counts its reads. */
function sequenceReader(states: Array<SetupState | null>) {
  const reader = {
    reads: 0,
    read: (): SetupState | null => states[Math.min(reader.reads++, states.length - 1)] ?? null,
  };
  return reader;
}

for (const c of [
  { name: 'completed', last: completedState(), expected: 'completed' },
  { name: 'cancelled', last: cancelledState(), expected: 'cancelled' },
] as const) {
  test(`waitForSetupCompletion resolves ${c.name} state after polling through earlier states`, async () => {
    const reader = sequenceReader([null, makeState(), c.last]);

    const result = await waitForSetupCompletion({
      readSetupState: reader.read,
      ...fakeClock(),
    });

    assert.equal(result, c.expected);
    assert.equal(reader.reads, 3);
  });
}

test('ensureLauncherSetupReady launches setup app and resumes only after completion', async () => {
  const calls: string[] = [];
  const reader = sequenceReader([
    null,
    makeState(),
    completedState({ pluginInstallStatus: 'installed', pluginInstallPathSummary: '/tmp/mpv' }),
  ]);

  const ready = await ensureLauncherSetupReady({
    readSetupState: reader.read,
    launchSetupApp: () => {
      calls.push('launch');
    },
    ...fakeClock(),
  });

  assert.equal(ready, true);
  assert.deepEqual(calls, ['launch']);
});

test('ensureLauncherSetupReady bypasses setup gate when external yomitan is configured', async () => {
  const calls: string[] = [];

  const ready = await ensureLauncherSetupReady({
    readSetupState: () => null,
    isExternalYomitanConfigured: () => true,
    launchSetupApp: () => {
      calls.push('launch');
    },
    ...fakeClock(),
  });

  assert.equal(ready, true);
  assert.deepEqual(calls, []);
});

test('ensureLauncherSetupReady waits for finish after legacy mpv plugin removal', async () => {
  const calls: string[] = [];
  let legacyPluginInstalled = true;
  // The completion timestamp moves forward once the setup app has finished.
  const reader = sequenceReader([
    completedState({ yomitanSetupMode: null, lastSeenYomitanDictionaryCount: 0 }),
    completedState({ yomitanSetupMode: null, lastSeenYomitanDictionaryCount: 0 }),
    completedState({
      completedAt: '2026-05-12T14:40:00.000Z',
      yomitanSetupMode: null,
      lastSeenYomitanDictionaryCount: 0,
    }),
  ]);

  const ready = await ensureLauncherSetupReady({
    readSetupState: reader.read,
    hasLegacyMpvPlugin: () => legacyPluginInstalled,
    launchSetupApp: () => {
      calls.push('launch');
      legacyPluginInstalled = false;
    },
    ...fakeClock(),
  });

  assert.equal(ready, true);
  assert.deepEqual(calls, ['launch']);
  assert.equal(reader.reads >= 3, true);
});

test('ensureLauncherSetupReady lets users continue without removing a legacy mpv plugin', async () => {
  const calls: string[] = [];
  const reader = sequenceReader([
    completedState({ lastSeenYomitanDictionaryCount: 2 }),
    completedState({ lastSeenYomitanDictionaryCount: 2 }),
    completedState({ completedAt: '2026-05-12T14:30:00.000Z', lastSeenYomitanDictionaryCount: 2 }),
  ]);

  const ready = await ensureLauncherSetupReady({
    readSetupState: reader.read,
    hasLegacyMpvPlugin: () => true,
    launchSetupApp: () => {
      calls.push('launch');
    },
    ...fakeClock(),
  });

  assert.equal(ready, true);
  assert.deepEqual(calls, ['launch']);
});

test('ensureLauncherSetupReady fails on timeout/cancelled state', async () => {
  const result = await ensureLauncherSetupReady({
    readSetupState: cancelledState,
    launchSetupApp: () => undefined,
    ...fakeClock(),
  });

  assert.equal(result, false);
});

test('ensureLauncherSetupReady ignores stale cancelled state after launching setup app', async () => {
  const reader = sequenceReader([
    cancelledState(),
    cancelledState(),
    makeState(),
    completedState({
      completionSource: 'legacy_auto_detected',
      pluginInstallStatus: 'installed',
      pluginInstallPathSummary: '/tmp/mpv',
    }),
  ]);

  const result = await ensureLauncherSetupReady({
    readSetupState: reader.read,
    launchSetupApp: () => undefined,
    ...fakeClock(),
  });

  assert.equal(result, true);
});

test('Hachidori setup ignores completed Yomitan state and external Yomitan profiles', async () => {
  let state: SetupState = { ...createDefaultSetupState(), status: 'completed' };
  let launched = 0;
  let polls = 0;
  const ready = await ensureLauncherSetupReady({
    dictionaryBackend: 'hachidori',
    readSetupState: () => state,
    isExternalYomitanConfigured: () => true,
    launchSetupApp: () => {
      launched += 1;
    },
    sleep: async () => {
      polls += 1;
      state = { ...state, dictionaryBackend: 'hachidori', lastSeenYomitanDictionaryCount: 1 };
    },
    now: () => polls,
    timeoutMs: 5,
    pollIntervalMs: 1,
  });
  assert.equal(ready, true);
  assert.equal(launched, 1);
  assert.equal(polls, 1);
});

test('matching Hachidori completion resumes playback without launching setup', async () => {
  const ready = await ensureLauncherSetupReady({
    dictionaryBackend: 'hachidori',
    readSetupState: () => ({
      ...createDefaultSetupState(),
      dictionaryBackend: 'hachidori',
      status: 'completed',
      lastSeenYomitanDictionaryCount: 1,
    }),
    launchSetupApp: () => assert.fail('setup should not open'),
    sleep: async () => undefined,
    now: () => 0,
    timeoutMs: 5,
    pollIntervalMs: 1,
  });
  assert.equal(ready, true);
});

test('a backend that finished setup earlier passes the gate after switching back', async () => {
  const ready = await ensureLauncherSetupReady({
    dictionaryBackend: 'yomitan',
    readSetupState: () => ({
      ...createDefaultSetupState(),
      dictionaryBackend: 'hachidori',
      status: 'incomplete',
      completedDictionaryBackends: ['yomitan'],
    }),
    launchSetupApp: () => assert.fail('setup should not open'),
    sleep: async () => undefined,
    now: () => 0,
    timeoutMs: 5,
    pollIntervalMs: 1,
  });
  assert.equal(ready, true);
});

test('gate follows the running app backend until it restarts into the configured one', async () => {
  const warnings: string[] = [];
  const state = {
    ...createDefaultSetupState(),
    dictionaryBackend: 'yomitan' as const,
    status: 'completed' as const,
  };
  assert.equal(
    await resolveLauncherGateBackend({
      configuredBackend: 'hachidori',
      state,
      isAppRunning: async () => true,
      warn: (message) => warnings.push(message),
    }),
    'yomitan',
  );
  assert.match(warnings[0] ?? '', /restart it to switch to hachidori/);
  assert.equal(
    await resolveLauncherGateBackend({
      configuredBackend: 'hachidori',
      state,
      isAppRunning: async () => false,
    }),
    'hachidori',
  );
  const ready = await ensureLauncherSetupReady({
    dictionaryBackend: 'hachidori',
    isAppRunning: async () => true,
    readSetupState: () => state,
    launchSetupApp: () => assert.fail('the running Yomitan app already completed setup'),
    sleep: async () => undefined,
    now: () => 0,
    timeoutMs: 5,
    pollIntervalMs: 1,
  });
  assert.equal(ready, true);
});
