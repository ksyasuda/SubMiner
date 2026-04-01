import assert from 'node:assert/strict';
import test from 'node:test';

import type { JellyfinRuntimeInput } from './jellyfin-runtime';
import { createJellyfinRuntime } from './jellyfin-runtime';

test('jellyfin runtime reuses existing setup window', () => {
  const calls: string[] = [];
  let windowCount = 0;

  const runtime = createJellyfinRuntime({
    getResolvedConfig: () =>
      ({
        jellyfin: {
          enabled: true,
          serverUrl: 'https://media.example',
          username: 'demo',
        },
      }) as ReturnType<JellyfinRuntimeInput['getResolvedConfig']>,
    getEnv: () => undefined,
    patchRawConfig: () => {},
    defaultJellyfinConfig: {
      enabled: false,
      serverUrl: '',
      username: '',
    } as JellyfinRuntimeInput['defaultJellyfinConfig'],
    tokenStore: {
      loadSession: () => null,
      saveSession: () => {},
      clearSession: () => {},
    },
    platform: 'linux',
    execPath: '/usr/bin/electron',
    defaultMpvLogPath: '/tmp/mpv.log',
    defaultMpvArgs: ['--idle=yes'],
    connectTimeoutMs: 1000,
    autoLaunchTimeoutMs: 1000,
    langPref: 'ja,en',
    progressIntervalMs: 3000,
    ticksPerSecond: 10_000_000,
    getMpvSocketPath: () => '/tmp/mpv.sock',
    getMpvClient: () => null,
    setMpvClient: () => {},
    createMpvClient: () => ({}),
    sendMpvCommand: () => {},
    applyJellyfinMpvDefaults: () => {},
    showMpvOsd: () => {},
    removeSocketPath: () => {},
    spawnMpv: () => ({}),
    wait: async () => {},
    authenticateWithPassword: async () => {
      throw new Error('not used');
    },
    listJellyfinLibraries: async () => [],
    listJellyfinItems: async () => [],
    listJellyfinSubtitleTracks: async () => [],
    writeJellyfinPreviewAuth: () => {},
    resolvePlaybackPlan: async () => ({}),
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    createRemoteSessionService: () => ({}),
    defaultDeviceId: 'device',
    defaultClientName: 'SubMiner',
    defaultClientVersion: '1.0.0',
    createBrowserWindow: () => {
      windowCount += 1;
      return {
        webContents: {
          on: () => {},
        },
        loadURL: () => {
          calls.push('loadURL');
        },
        on: () => {},
        focus: () => {
          calls.push('focus');
        },
        close: () => {},
        isDestroyed: () => false,
      };
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
    logInfo: () => {},
    logWarn: () => {},
    logDebug: () => {},
    logError: () => {},
  });

  runtime.openJellyfinSetupWindow();
  runtime.openJellyfinSetupWindow();

  assert.equal(windowCount, 1);
  assert.deepEqual(calls, ['loadURL', 'focus']);
  assert.equal(runtime.getQuitOnDisconnectArmed(), false);
  assert.ok(runtime.getSetupWindow());
});
