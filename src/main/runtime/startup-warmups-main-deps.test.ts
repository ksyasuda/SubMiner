import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildLaunchBackgroundWarmupTaskMainDepsHandler,
  createBuildStartBackgroundWarmupsMainDepsHandler,
} from './startup-warmups-main-deps';

test('startup warmups main deps builders map callbacks', async () => {
  const calls: string[] = [];

  const launch = createBuildLaunchBackgroundWarmupTaskMainDepsHandler({
    now: () => 11,
    logDebug: (message) => calls.push(`debug:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
  })();
  assert.equal(launch.now(), 11);
  launch.logDebug('x');
  launch.logWarn('y');

  const start = createBuildStartBackgroundWarmupsMainDepsHandler({
    getStarted: () => false,
    setStarted: (started) => calls.push(`started:${started}`),
    isTexthookerOnlyMode: () => false,
    launchTask: (label, task) => {
      calls.push(`launch:${label}`);
      void task();
    },
    createMecabTokenizerAndCheck: async () => {
      calls.push('mecab');
    },
    ensureYomitanExtensionLoaded: async () => {
      calls.push('yomitan');
    },
    prewarmSubtitleDictionaries: async () => {
      calls.push('dict');
    },
    shouldWarmupMecab: () => false,
    shouldWarmupYomitanExtension: () => true,
    shouldWarmupSubtitleDictionaries: () => false,
    shouldWarmupJellyfinRemoteSession: () => true,
    shouldAutoConnectJellyfinRemote: () => true,
    startJellyfinRemoteSession: async () => {
      calls.push('jellyfin');
    },
  })();
  assert.equal(start.getStarted(), false);
  start.setStarted(true);
  assert.equal(start.isTexthookerOnlyMode(), false);
  start.launchTask('demo', async () => {
    calls.push('task');
  });
  await start.createMecabTokenizerAndCheck();
  await start.ensureYomitanExtensionLoaded();
  await start.prewarmSubtitleDictionaries();
  assert.equal(start.shouldWarmupMecab(), false);
  assert.equal(start.shouldWarmupYomitanExtension(), true);
  assert.equal(start.shouldWarmupSubtitleDictionaries(), false);
  assert.equal(start.shouldWarmupJellyfinRemoteSession(), true);
  assert.equal(start.shouldAutoConnectJellyfinRemote(), true);
  await start.startJellyfinRemoteSession();

  assert.deepEqual(calls, [
    'debug:x',
    'warn:y',
    'started:true',
    'launch:demo',
    'task',
    'mecab',
    'yomitan',
    'dict',
    'jellyfin',
  ]);
});
