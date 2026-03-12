import test from 'node:test';
import assert from 'node:assert/strict';
import { ensureLauncherSetupReady, waitForSetupCompletion } from './setup-gate';
import type { SetupState } from '../src/shared/setup-state';

test('waitForSetupCompletion resolves completed and cancelled states', async () => {
  const sequence: Array<SetupState | null> = [
    null,
    {
      version: 3,
      status: 'in_progress',
      completedAt: null,
      completionSource: null,
      yomitanSetupMode: null,
      lastSeenYomitanDictionaryCount: 0,
      pluginInstallStatus: 'unknown',
      pluginInstallPathSummary: null,
      windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
      windowsMpvShortcutLastStatus: 'unknown',
    },
    {
      version: 3,
      status: 'completed',
      completedAt: '2026-03-07T00:00:00.000Z',
      completionSource: 'user',
      yomitanSetupMode: 'internal',
      lastSeenYomitanDictionaryCount: 1,
      pluginInstallStatus: 'skipped',
      pluginInstallPathSummary: null,
      windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
      windowsMpvShortcutLastStatus: 'skipped',
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
          version: 3,
          status: 'in_progress',
          completedAt: null,
          completionSource: null,
          yomitanSetupMode: null,
          lastSeenYomitanDictionaryCount: 0,
          pluginInstallStatus: 'unknown',
          pluginInstallPathSummary: null,
          windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
          windowsMpvShortcutLastStatus: 'unknown',
        };
      }
      return {
        version: 3,
        status: 'completed',
        completedAt: '2026-03-07T00:00:00.000Z',
        completionSource: 'user',
        yomitanSetupMode: 'internal',
        lastSeenYomitanDictionaryCount: 1,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: '/tmp/mpv',
        windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
        windowsMpvShortcutLastStatus: 'installed',
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

test('ensureLauncherSetupReady fails on timeout/cancelled state', async () => {
  const result = await ensureLauncherSetupReady({
    readSetupState: () => ({
      version: 3,
      status: 'cancelled',
      completedAt: null,
      completionSource: null,
      yomitanSetupMode: null,
      lastSeenYomitanDictionaryCount: 0,
      pluginInstallStatus: 'unknown',
      pluginInstallPathSummary: null,
      windowsMpvShortcutPreferences: { startMenuEnabled: true, desktopEnabled: true },
      windowsMpvShortcutLastStatus: 'unknown',
    }),
    launchSetupApp: () => undefined,
    sleep: async () => undefined,
    now: () => 0,
    timeoutMs: 5_000,
    pollIntervalMs: 100,
  });

  assert.equal(result, false);
});
