import assert from 'node:assert/strict';
import test from 'node:test';

import { createDictionarySupportRuntime } from './dictionary-support-runtime';

function createRuntime() {
  const state = {
    currentMediaPath: null as string | null,
    currentMediaTitle: null as string | null,
    jlptLookupSet: 0,
    frequencyLookupSet: 0,
    trackerCalls: [] as Array<{ path: string; title: string | null }>,
    characterDictionaryConfig: {
      enabled: false,
      maxLoaded: 1,
      profileScope: 'global',
    },
    youtubePlaybackActive: false,
  };

  const runtime = createDictionarySupportRuntime({
    platform: 'darwin',
    dirname: '/repo/dist/main',
    appPath: '/Applications/SubMiner.app/Contents/Resources/app.asar',
    resourcesPath: '/Applications/SubMiner.app/Contents/Resources',
    userDataPath: '/Users/a/Library/Application Support/SubMiner',
    appUserDataPath: '/Users/a/Library/Application Support/SubMiner',
    homeDir: '/Users/a',
    cwd: '/repo',
    subtitlePositionsDir: '/Users/a/Library/Application Support/SubMiner/subtitle-positions',
    getResolvedConfig: () =>
      ({
        subtitleStyle: {
          enableJlpt: false,
          frequencyDictionary: {
            enabled: false,
            sourcePath: '',
          },
        },
        anilist: {
          characterDictionary: {
            enabled: false,
            maxLoaded: 1,
            profileScope: 'global',
            collapsibleSections: {
              description: false,
              glossary: false,
              termEntry: false,
              nameReading: false,
            },
          },
        },
        ankiConnect: {
          behavior: {
            notificationType: 'none',
          },
        },
      }) as never,
    isJlptEnabled: () => false,
    isFrequencyDictionaryEnabled: () => false,
    getFrequencyDictionarySourcePath: () => undefined,
    setJlptLevelLookup: () => {
      state.jlptLookupSet += 1;
    },
    setFrequencyRankLookup: () => {
      state.frequencyLookupSet += 1;
    },
    logInfo: () => {},
    logDebug: () => {},
    logWarn: () => {},
    isRemoteMediaPath: (mediaPath) => mediaPath.startsWith('remote:'),
    getCurrentMediaPath: () => state.currentMediaPath,
    setCurrentMediaPath: (mediaPath) => {
      state.currentMediaPath = mediaPath;
    },
    getCurrentMediaTitle: () => state.currentMediaTitle,
    setCurrentMediaTitle: (title) => {
      state.currentMediaTitle = title;
    },
    getPendingSubtitlePosition: () => null,
    loadSubtitlePosition: () => null,
    clearPendingSubtitlePosition: () => {},
    setSubtitlePosition: () => {},
    broadcastSubtitlePosition: () => {},
    broadcastToOverlayWindows: () => {},
    getTracker: () =>
      ({
        handleMediaChange: (path: string, title: string | null) => {
          state.trackerCalls.push({ path, title });
        },
      }) as never,
    getMpvClient: () => null,
    defaultImmersionDbPath: '/tmp/immersion.db',
    guessAnilistMediaInfo: async () => null,
    getCollapsibleSectionOpenState: () => false,
    isCharacterDictionaryEnabled: () => state.characterDictionaryConfig.enabled,
    isYoutubePlaybackActiveNow: () => state.youtubePlaybackActive,
    waitForYomitanMutationReady: async () => {},
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: async () => false,
    deleteYomitanDictionary: async () => false,
    upsertYomitanDictionarySettings: async () => false,
    getCharacterDictionaryConfig: () => state.characterDictionaryConfig as never,
    notifyCharacterDictionaryAutoSyncStatus: () => {},
    characterDictionaryAutoSyncCompleteDeps: {
      hasParserWindow: () => false,
      clearParserCaches: () => {},
      invalidateTokenizationCache: () => {},
      refreshSubtitlePrefetch: () => {},
      refreshCurrentSubtitle: () => {},
      logInfo: () => {},
    },
    getMainWindow: () => null,
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    getRestoreVisibleOverlayOnModalClose: () => new Set<string>(),
    sendToActiveOverlayWindow: () => true,
  });

  return { runtime, state };
}

test('dictionary support runtime wires field grouping resolver and callback', async () => {
  const { runtime } = createRuntime();
  const choice = {
    keepNoteId: 1,
    deleteNoteId: 2,
    deleteDuplicate: false,
    cancelled: false,
  };

  const callback = runtime.createFieldGroupingCallback();
  const pending = callback({} as never);
  const resolver = runtime.getFieldGroupingResolver();
  assert.ok(resolver);
  resolver(choice as never);
  assert.deepEqual(await pending, choice);
  assert.equal(typeof runtime.getFieldGroupingResolver(), 'function');

  runtime.setFieldGroupingResolver(null);
  assert.equal(runtime.getFieldGroupingResolver(), null);
});

test('dictionary support runtime resolves media paths and keeps title in sync', () => {
  const { runtime, state } = createRuntime();

  runtime.updateCurrentMediaTitle('  Example Title  ');
  runtime.updateCurrentMediaPath('remote://media' as never);
  assert.equal(state.currentMediaTitle, 'Example Title');
  assert.equal(runtime.resolveMediaPathForJimaku('remote://media'), 'Example Title');

  runtime.updateCurrentMediaPath('local.mp4' as never);
  assert.equal(state.currentMediaTitle, null);
  assert.equal(runtime.resolveMediaPathForJimaku('remote://media'), 'remote://media');
});

test('dictionary support runtime skips disabled lookup and sync paths', async () => {
  const { runtime, state } = createRuntime();

  await runtime.ensureJlptDictionaryLookup();
  await runtime.ensureFrequencyDictionaryLookup();
  runtime.scheduleCharacterDictionarySync();

  assert.equal(state.jlptLookupSet, 0);
  assert.equal(state.frequencyLookupSet, 0);
});

test('dictionary support runtime syncs immersion media from current state', async () => {
  const { runtime, state } = createRuntime();

  runtime.updateCurrentMediaTitle(' Example Title ');
  runtime.updateCurrentMediaPath('remote://media' as never);
  await runtime.seedImmersionMediaFromCurrentMedia();
  runtime.syncImmersionMediaState();

  assert.deepEqual(state.trackerCalls, [
    { path: 'remote://media', title: 'Example Title' },
    { path: 'remote://media', title: 'Example Title' },
  ]);
});

test('dictionary support runtime gates character dictionary auto-sync scheduling', () => {
  const { runtime, state } = createRuntime();
  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  let timeoutCalls = 0;

  try {
    globalThis.setTimeout = ((handler: TimerHandler, timeout?: number, ...args: never[]) => {
      timeoutCalls += 1;
      return originalSetTimeout(handler, timeout ?? 0, ...args);
    }) as typeof globalThis.setTimeout;
    globalThis.clearTimeout = ((handle: number | NodeJS.Timeout | undefined) => {
      originalClearTimeout(handle);
    }) as typeof globalThis.clearTimeout;

    runtime.scheduleCharacterDictionarySync();
    assert.equal(timeoutCalls, 0);

    state.characterDictionaryConfig = {
      enabled: true,
      maxLoaded: 1,
      profileScope: 'global',
    };
    runtime.scheduleCharacterDictionarySync();
    assert.equal(timeoutCalls, 1);

    state.youtubePlaybackActive = true;
    runtime.scheduleCharacterDictionarySync();
    assert.equal(timeoutCalls, 1);
  } finally {
    globalThis.setTimeout = originalSetTimeout;
    globalThis.clearTimeout = originalClearTimeout;
  }
});
