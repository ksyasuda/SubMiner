import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createLaunchBackgroundWarmupTaskHandler,
  createStartBackgroundWarmupsHandler,
} from './startup-warmups';

test('launchBackgroundWarmupTask logs completion timing', async () => {
  const debugLogs: string[] = [];
  const launchTask = createLaunchBackgroundWarmupTaskHandler({
    now: (() => {
      let tick = 0;
      return () => ++tick * 10;
    })(),
    logDebug: (message) => debugLogs.push(message),
    logWarn: () => {},
  });

  launchTask('demo', async () => {});
  await Promise.resolve();
  assert.ok(debugLogs.some((line) => line.includes('[startup-warmup] demo completed in')));
});

test('startBackgroundWarmups no-ops when already started', () => {
  let launches = 0;
  const startWarmups = createStartBackgroundWarmupsHandler({
    getStarted: () => true,
    setStarted: () => {},
    isTexthookerOnlyMode: () => false,
    launchTask: () => {
      launches += 1;
    },
    createMecabTokenizerAndCheck: async () => {},
    ensureYomitanExtensionLoaded: async () => {},
    prewarmSubtitleDictionaries: async () => {},
    shouldAutoConnectJellyfinRemote: () => false,
    startJellyfinRemoteSession: async () => {},
  });

  startWarmups();
  assert.equal(launches, 0);
});

test('startBackgroundWarmups schedules base warmups and optional jellyfin warmup', () => {
  const labels: string[] = [];
  let started = false;
  const startWarmups = createStartBackgroundWarmupsHandler({
    getStarted: () => started,
    setStarted: (value) => {
      started = value;
    },
    isTexthookerOnlyMode: () => false,
    launchTask: (label) => {
      labels.push(label);
    },
    createMecabTokenizerAndCheck: async () => {},
    ensureYomitanExtensionLoaded: async () => {},
    prewarmSubtitleDictionaries: async () => {},
    shouldAutoConnectJellyfinRemote: () => true,
    startJellyfinRemoteSession: async () => {},
  });

  startWarmups();
  assert.equal(started, true);
  assert.deepEqual(labels, [
    'mecab',
    'yomitan-extension',
    'subtitle-dictionaries',
    'jellyfin-remote-session',
  ]);
});
