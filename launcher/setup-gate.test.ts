import test from 'node:test';
import assert from 'node:assert/strict';
import {
  ensureLauncherSetupReady,
  resolveLauncherGateBackend,
  waitForSetupCompletion,
} from './setup-gate';
import { createDefaultSetupState, type SetupState } from '../src/shared/setup-state';

const commandLineSetupDefaults = {
  bunInstallStatus: 'unknown',
  launcherInstallStatus: 'unknown',
  launcherInstallPath: null,
} satisfies Pick<SetupState, 'bunInstallStatus' | 'launcherInstallStatus' | 'launcherInstallPath'>;

test('waitForSetupCompletion resolves completed and cancelled states', async () => {
  const sequence: Array<SetupState | null> = [
    null,
    {
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
      ...commandLineSetupDefaults,
    },
    {
      version: 4,
      status: 'completed',
      completedAt: '2026-03-07T00:00:00.000Z',
      completionSource: 'user',
      yomitanSetupMode: 'internal',
      lastSeenYomitanDictionaryCount: 1,
      pluginInstallStatus: 'skipped',
      pluginInstallPathSummary: null,
      windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
      windowsMpvShortcutLastStatus: 'skipped',
      ...commandLineSetupDefaults,
    },
  ];

  const result = await waitForSetupCompletion({
    readSetupState: () => sequence.shift() ?? null,
    sleep: async () => undefined,
    now: (() => {
      let value = 0;
      return () => (value += 100);
    })(),
    timeoutMs: 5_000,
    pollIntervalMs: 100,
  });

  assert.equal(result, 'completed');
});

test('ensureLauncherSetupReady launches setup app and resumes only after completion', async () => {
  const calls: string[] = [];
  let reads = 0;

  const ready = await ensureLauncherSetupReady({
    readSetupState: () => {
      reads += 1;
      if (reads === 1) return null;
      if (reads === 2) {
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
          ...commandLineSetupDefaults,
        };
      }
      return {
        version: 4,
        status: 'completed',
        completedAt: '2026-03-07T00:00:00.000Z',
        completionSource: 'user',
        yomitanSetupMode: 'internal',
        lastSeenYomitanDictionaryCount: 1,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: '/tmp/mpv',
        windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
        windowsMpvShortcutLastStatus: 'installed',
        ...commandLineSetupDefaults,
      };
    },
    launchSetupApp: () => {
      calls.push('launch');
    },
    sleep: async () => undefined,
    now: (() => {
      let value = 0;
      return () => (value += 100);
    })(),
    timeoutMs: 5_000,
    pollIntervalMs: 100,
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
    sleep: async () => undefined,
    now: () => 0,
    timeoutMs: 5_000,
    pollIntervalMs: 100,
  });

  assert.equal(ready, true);
  assert.deepEqual(calls, []);
});

test('ensureLauncherSetupReady waits for finish after legacy mpv plugin removal', async () => {
  const calls: string[] = [];
  let legacyPluginInstalled = true;
  let reads = 0;

  const ready = await ensureLauncherSetupReady({
    readSetupState: () => {
      reads += 1;
      return {
        version: 4,
        status: 'completed',
        completedAt: reads < 3 ? '2026-03-07T00:00:00.000Z' : '2026-05-12T14:40:00.000Z',
        completionSource: 'user',
        yomitanSetupMode: null,
        lastSeenYomitanDictionaryCount: 0,
        pluginInstallStatus: 'unknown',
        pluginInstallPathSummary: null,
        windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
        windowsMpvShortcutLastStatus: 'unknown',
        ...commandLineSetupDefaults,
      };
    },
    hasLegacyMpvPlugin: () => legacyPluginInstalled,
    launchSetupApp: () => {
      calls.push('launch');
      legacyPluginInstalled = false;
    },
    sleep: async () => undefined,
    now: (() => {
      let value = 0;
      return () => (value += 100);
    })(),
    timeoutMs: 5_000,
    pollIntervalMs: 100,
  });

  assert.equal(ready, true);
  assert.deepEqual(calls, ['launch']);
  assert.equal(reads >= 3, true);
});

test('ensureLauncherSetupReady lets users continue without removing a legacy mpv plugin', async () => {
  const calls: string[] = [];
  let reads = 0;

  const ready = await ensureLauncherSetupReady({
    readSetupState: () => {
      reads += 1;
      return {
        version: 4,
        status: 'completed',
        completedAt: reads < 3 ? '2026-03-07T00:00:00.000Z' : '2026-05-12T14:30:00.000Z',
        completionSource: 'user',
        yomitanSetupMode: 'internal',
        lastSeenYomitanDictionaryCount: 2,
        pluginInstallStatus: 'unknown',
        pluginInstallPathSummary: null,
        windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
        windowsMpvShortcutLastStatus: 'unknown',
        ...commandLineSetupDefaults,
      };
    },
    hasLegacyMpvPlugin: () => true,
    launchSetupApp: () => {
      calls.push('launch');
    },
    sleep: async () => undefined,
    now: (() => {
      let value = 0;
      return () => (value += 100);
    })(),
    timeoutMs: 5_000,
    pollIntervalMs: 100,
  });

  assert.equal(ready, true);
  assert.deepEqual(calls, ['launch']);
});

test('ensureLauncherSetupReady fails on timeout/cancelled state', async () => {
  const result = await ensureLauncherSetupReady({
    readSetupState: () => ({
      version: 4,
      status: 'cancelled',
      completedAt: null,
      completionSource: null,
      yomitanSetupMode: null,
      lastSeenYomitanDictionaryCount: 0,
      pluginInstallStatus: 'unknown',
      pluginInstallPathSummary: null,
      windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
      windowsMpvShortcutLastStatus: 'unknown',
      ...commandLineSetupDefaults,
    }),
    launchSetupApp: () => undefined,
    sleep: async () => undefined,
    now: (() => {
      let value = 0;
      return () => (value += 100);
    })(),
    timeoutMs: 5_000,
    pollIntervalMs: 100,
  });

  assert.equal(result, false);
});

test('ensureLauncherSetupReady ignores stale cancelled state after launching setup app', async () => {
  let reads = 0;

  const result = await ensureLauncherSetupReady({
    readSetupState: () => {
      reads += 1;
      if (reads <= 2) {
        return {
          version: 4,
          status: 'cancelled',
          completedAt: null,
          completionSource: null,
          yomitanSetupMode: null,
          lastSeenYomitanDictionaryCount: 0,
          pluginInstallStatus: 'unknown',
          pluginInstallPathSummary: null,
          windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
          windowsMpvShortcutLastStatus: 'unknown',
          ...commandLineSetupDefaults,
        };
      }
      if (reads === 3) {
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
          ...commandLineSetupDefaults,
        };
      }
      return {
        version: 4,
        status: 'completed',
        completedAt: '2026-03-07T00:00:00.000Z',
        completionSource: 'legacy_auto_detected',
        yomitanSetupMode: 'internal',
        lastSeenYomitanDictionaryCount: 1,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: '/tmp/mpv',
        windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
        windowsMpvShortcutLastStatus: 'unknown',
        ...commandLineSetupDefaults,
      };
    },
    launchSetupApp: () => undefined,
    sleep: async () => undefined,
    now: (() => {
      let value = 0;
      return () => (value += 100);
    })(),
    timeoutMs: 5_000,
    pollIntervalMs: 100,
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
