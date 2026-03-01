import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createLaunchBackgroundWarmupTaskHandler,
  createStartBackgroundWarmupsHandler,
} from './startup-warmups';

function shouldAutoConnectJellyfinRemote(config: {
  enabled: boolean;
  remoteControlEnabled: boolean;
  remoteControlAutoConnect: boolean;
}): boolean {
  return config.enabled && config.remoteControlEnabled && config.remoteControlAutoConnect;
}

function createDeferred(): {
  promise: Promise<void>;
  resolve: () => void;
} {
  let resolve!: () => void;
  const promise = new Promise<void>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

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
    shouldWarmupMecab: () => true,
    shouldWarmupYomitanExtension: () => true,
    shouldWarmupSubtitleDictionaries: () => true,
    shouldWarmupJellyfinRemoteSession: () => true,
    shouldAutoConnectJellyfinRemote: () => false,
    startJellyfinRemoteSession: async () => {},
  });

  startWarmups();
  assert.equal(launches, 0);
});

test('startBackgroundWarmups respects per-integration warmup toggles', () => {
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
    shouldWarmupMecab: () => false,
    shouldWarmupYomitanExtension: () => true,
    shouldWarmupSubtitleDictionaries: () => false,
    shouldWarmupJellyfinRemoteSession: () => false,
    shouldAutoConnectJellyfinRemote: () =>
      shouldAutoConnectJellyfinRemote({
        enabled: true,
        remoteControlEnabled: true,
        remoteControlAutoConnect: true,
      }),
    startJellyfinRemoteSession: async () => {},
  });

  startWarmups();
  assert.equal(started, true);
  assert.deepEqual(labels, ['subtitle-tokenization']);
});

test('startBackgroundWarmups schedules jellyfin warmup when all jellyfin flags are enabled', () => {
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
    shouldWarmupMecab: () => true,
    shouldWarmupYomitanExtension: () => true,
    shouldWarmupSubtitleDictionaries: () => true,
    shouldWarmupJellyfinRemoteSession: () => true,
    shouldAutoConnectJellyfinRemote: () =>
      shouldAutoConnectJellyfinRemote({
        enabled: true,
        remoteControlEnabled: true,
        remoteControlAutoConnect: true,
      }),
    startJellyfinRemoteSession: async () => {},
  });

  startWarmups();
  assert.equal(started, true);
  assert.deepEqual(labels, ['subtitle-tokenization', 'jellyfin-remote-session']);
});

test('startBackgroundWarmups skips jellyfin warmup when warmup is deferred', () => {
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
    shouldWarmupMecab: () => false,
    shouldWarmupYomitanExtension: () => true,
    shouldWarmupSubtitleDictionaries: () => false,
    shouldWarmupJellyfinRemoteSession: () => false,
    shouldAutoConnectJellyfinRemote: () =>
      shouldAutoConnectJellyfinRemote({
        enabled: true,
        remoteControlEnabled: true,
        remoteControlAutoConnect: true,
      }),
    startJellyfinRemoteSession: async () => {},
  });

  startWarmups();
  assert.equal(started, true);
  assert.deepEqual(labels, ['subtitle-tokenization']);
});

test('startBackgroundWarmups logs per-stage progress for enabled tokenization warmups', async () => {
  const debugLogs: string[] = [];
  const labels: string[] = [];
  let started = false;
  const startWarmups = createStartBackgroundWarmupsHandler({
    getStarted: () => started,
    setStarted: (value) => {
      started = value;
    },
    isTexthookerOnlyMode: () => false,
    launchTask: (label, task) => {
      labels.push(label);
      void task();
    },
    createMecabTokenizerAndCheck: async () => {},
    ensureYomitanExtensionLoaded: async () => {},
    prewarmSubtitleDictionaries: async () => {},
    shouldWarmupMecab: () => true,
    shouldWarmupYomitanExtension: () => true,
    shouldWarmupSubtitleDictionaries: () => true,
    shouldWarmupJellyfinRemoteSession: () => true,
    shouldAutoConnectJellyfinRemote: () => true,
    startJellyfinRemoteSession: async () => {},
    logDebug: (message) => {
      debugLogs.push(message);
    },
  });

  startWarmups();
  await Promise.resolve();
  await Promise.resolve();

  assert.deepEqual(labels, ['subtitle-tokenization', 'jellyfin-remote-session']);
  assert.ok(debugLogs.includes('[startup-warmup] stage start: yomitan-extension'));
  assert.ok(debugLogs.includes('[startup-warmup] stage ready: yomitan-extension'));
  assert.ok(debugLogs.includes('[startup-warmup] stage start: mecab'));
  assert.ok(debugLogs.includes('[startup-warmup] stage ready: mecab'));
  assert.ok(debugLogs.includes('[startup-warmup] stage start: subtitle-dictionaries'));
  assert.ok(debugLogs.includes('[startup-warmup] stage ready: subtitle-dictionaries'));
  assert.ok(debugLogs.includes('[startup-warmup] stage start: jellyfin-remote-session'));
  assert.ok(debugLogs.includes('[startup-warmup] stage ready: jellyfin-remote-session'));
});

test('startBackgroundWarmups starts mecab and dictionary warmups without waiting for yomitan warmup', async () => {
  const startedStages: string[] = [];
  let started = false;
  let subtitleTokenizationTask: Promise<void> | null = null;
  const yomitanDeferred = createDeferred();
  const mecabDeferred = createDeferred();
  const subtitleDictionariesDeferred = createDeferred();

  const startWarmups = createStartBackgroundWarmupsHandler({
    getStarted: () => started,
    setStarted: (value) => {
      started = value;
    },
    isTexthookerOnlyMode: () => false,
    launchTask: (label, task) => {
      if (label === 'subtitle-tokenization') {
        subtitleTokenizationTask = task();
      }
    },
    createMecabTokenizerAndCheck: async () => {
      startedStages.push('mecab');
      await mecabDeferred.promise;
    },
    ensureYomitanExtensionLoaded: async () => {
      startedStages.push('yomitan-extension');
      await yomitanDeferred.promise;
    },
    prewarmSubtitleDictionaries: async () => {
      startedStages.push('subtitle-dictionaries');
      await subtitleDictionariesDeferred.promise;
    },
    shouldWarmupMecab: () => true,
    shouldWarmupYomitanExtension: () => true,
    shouldWarmupSubtitleDictionaries: () => true,
    shouldWarmupJellyfinRemoteSession: () => false,
    shouldAutoConnectJellyfinRemote: () => false,
    startJellyfinRemoteSession: async () => {},
  });

  startWarmups();
  await Promise.resolve();
  await Promise.resolve();

  assert.ok(subtitleTokenizationTask);
  assert.equal(startedStages.includes('yomitan-extension'), true);
  assert.equal(startedStages.includes('mecab'), true);
  assert.equal(startedStages.includes('subtitle-dictionaries'), true);

  yomitanDeferred.resolve();
  mecabDeferred.resolve();
  subtitleDictionariesDeferred.resolve();
  if (!subtitleTokenizationTask) {
    throw new Error('Expected subtitle tokenization warmup task');
  }
  await subtitleTokenizationTask;
});
