import assert from 'node:assert/strict';
import test from 'node:test';

import { createStartupSequenceRuntime } from './startup-sequence-runtime';

test('startup sequence delegates non-refresh headless command to initial args handler', async () => {
  const calls: string[] = [];

  const runtime = createStartupSequenceRuntime({
    appState: {
      initialArgs: { refreshKnownWords: false } as never,
      runtimeOptionsManager: null,
    },
    userDataPath: '/tmp/subminer',
    getResolvedConfig: () => ({ ankiConnect: { enabled: true } }) as never,
    anilist: {
      refreshAnilistClientSecretStateIfEnabled: async () => undefined,
      refreshRetryQueueState: () => {},
    },
    actions: {
      initializeDiscordPresenceService: async () => {},
      requestAppQuit: () => {},
    },
    logger: {
      error: () => {},
    },
    runHeadlessKnownWordRefresh: async () => {
      calls.push('refreshKnownWords');
    },
  });

  await runtime.runHeadlessInitialCommand({
    handleInitialArgs: () => {
      calls.push('handleInitialArgs');
    },
  });

  assert.deepEqual(calls, ['handleInitialArgs']);
});

test('startup sequence runs headless known-word refresh when requested', async () => {
  const calls: string[] = [];
  const runtimeOptionsManager = {
    getEffectiveAnkiConnectConfig: (config: never) => config,
  } as never;

  const runtime = createStartupSequenceRuntime({
    appState: {
      initialArgs: { refreshKnownWords: true } as never,
      runtimeOptionsManager,
    },
    userDataPath: '/tmp/subminer',
    getResolvedConfig: () => ({ ankiConnect: { enabled: true } }) as never,
    anilist: {
      refreshAnilistClientSecretStateIfEnabled: async () => undefined,
      refreshRetryQueueState: () => {},
    },
    actions: {
      initializeDiscordPresenceService: async () => {},
      requestAppQuit: () => {
        calls.push('requestAppQuit');
      },
    },
    logger: {
      error: () => {},
    },
    runHeadlessKnownWordRefresh: async (input) => {
      calls.push(`refresh:${input.userDataPath}`);
      assert.equal(input.runtimeOptionsManager, runtimeOptionsManager);
    },
  });

  await runtime.runHeadlessInitialCommand({
    handleInitialArgs: () => {
      calls.push('handleInitialArgs');
    },
  });

  assert.deepEqual(calls, ['refresh:/tmp/subminer']);
});

test('startup sequence runs deferred AniList and Discord init only for full startup', async () => {
  const calls: string[] = [];

  const runtime = createStartupSequenceRuntime({
    appState: {
      initialArgs: null,
      runtimeOptionsManager: null,
    },
    userDataPath: '/tmp/subminer',
    getResolvedConfig: () => ({ anilist: { enabled: true } }) as never,
    anilist: {
      refreshAnilistClientSecretStateIfEnabled: async (options) => {
        calls.push(`anilist:${options.force}:${options.allowSetupPrompt}`);
      },
      refreshRetryQueueState: () => {
        calls.push('retryQueue');
      },
    },
    actions: {
      initializeDiscordPresenceService: async () => {
        calls.push('discord');
      },
      requestAppQuit: () => {},
    },
    logger: {
      error: () => {},
    },
  });

  runtime.runPostStartupInitialization();
  await new Promise((resolve) => setImmediate(resolve));

  assert.deepEqual(calls, ['anilist:true:false', 'retryQueue', 'discord']);
});

test('startup sequence skips deferred startup side effects in minimal mode', async () => {
  const calls: string[] = [];

  const runtime = createStartupSequenceRuntime({
    appState: {
      initialArgs: { background: true } as never,
      runtimeOptionsManager: null,
    },
    userDataPath: '/tmp/subminer',
    getResolvedConfig: () => ({ anilist: { enabled: true } }) as never,
    anilist: {
      refreshAnilistClientSecretStateIfEnabled: async () => {
        calls.push('anilist');
      },
      refreshRetryQueueState: () => {
        calls.push('retryQueue');
      },
    },
    actions: {
      initializeDiscordPresenceService: async () => {
        calls.push('discord');
      },
      requestAppQuit: () => {},
    },
    logger: {
      error: () => {},
    },
    getStartupModeFlags: () => ({
      shouldUseMinimalStartup: true,
      shouldSkipHeavyStartup: false,
    }),
    isAnilistTrackingEnabled: () => true,
  });

  runtime.runPostStartupInitialization();
  await new Promise((resolve) => setImmediate(resolve));

  assert.deepEqual(calls, []);
});
