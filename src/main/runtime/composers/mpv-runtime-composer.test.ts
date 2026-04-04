import assert from 'node:assert/strict';
import test from 'node:test';
import type { MpvSubtitleRenderMetrics } from '../../../types';
import { composeMpvRuntimeHandlers } from './mpv-runtime-composer';

const BASE_METRICS: MpvSubtitleRenderMetrics = {
  subPos: 100,
  subFontSize: 36,
  subScale: 1,
  subMarginY: 0,
  subMarginX: 0,
  subFont: '',
  subSpacing: 0,
  subBold: false,
  subItalic: false,
  subBorderSize: 0,
  subShadowOffset: 0,
  subAssOverride: 'yes',
  subScaleByWindow: true,
  subUseMargins: true,
  osdHeight: 0,
  osdDimensions: null,
};

function createDeferred(): { promise: Promise<void>; resolve: () => void } {
  let resolve!: () => void;
  const promise = new Promise<void>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

class DefaultFakeMpvClient {
  connect(): void {}
  on(): void {}
}

function createDefaultMpvFixture() {
  return {
    bindMpvMainEventHandlersMainDeps: {
      appState: {
        initialArgs: null,
        overlayRuntimeInitialized: true,
        mpvClient: null,
        immersionTracker: null,
        subtitleTimingTracker: null,
        currentSubText: '',
        currentSubAssText: '',
        playbackPaused: null,
        previousSecondarySubVisibility: null,
      },
      getQuitOnDisconnectArmed: () => false,
      scheduleQuitCheck: () => {},
      quitApp: () => {},
      reportJellyfinRemoteStopped: () => {},
      syncOverlayMpvSubtitleSuppression: () => {},
      maybeRunAnilistPostWatchUpdate: async () => {},
      logSubtitleTimingError: () => {},
      broadcastToOverlayWindows: () => {},
      onSubtitleChange: () => {},
      refreshDiscordPresence: () => {},
      ensureImmersionTrackerInitialized: () => {},
      updateCurrentMediaPath: () => {},
      restoreMpvSubVisibility: () => {},
      getCurrentAnilistMediaKey: () => null,
      resetAnilistMediaTracking: () => {},
      maybeProbeAnilistDuration: () => {},
      ensureAnilistMediaGuess: () => {},
      syncImmersionMediaState: () => {},
      updateCurrentMediaTitle: () => {},
      resetAnilistMediaGuessState: () => {},
      reportJellyfinRemoteProgress: () => {},
      updateSubtitleRenderMetrics: () => {},
    },
    mpvClientRuntimeServiceFactoryMainDeps: {
      createClient: DefaultFakeMpvClient,
      getSocketPath: () => '/tmp/mpv.sock',
      getResolvedConfig: () => ({ auto_start_overlay: false }),
      isAutoStartOverlayEnabled: () => false,
      setOverlayVisible: () => {},
      isVisibleOverlayVisible: () => false,
      getReconnectTimer: () => null,
      setReconnectTimer: () => {},
    },
    updateMpvSubtitleRenderMetricsMainDeps: {
      getCurrentMetrics: () => BASE_METRICS,
      setCurrentMetrics: () => {},
      applyPatch: (
        current: MpvSubtitleRenderMetrics,
        patch: Partial<MpvSubtitleRenderMetrics>,
      ) => ({
        next: { ...current, ...patch },
        changed: true,
      }),
      broadcastMetrics: () => {},
    },
    tokenizer: {
      buildTokenizerDepsMainDeps: {
        getYomitanExt: () => null,
        getYomitanParserWindow: () => null,
        setYomitanParserWindow: () => {},
        getYomitanParserReadyPromise: () => null,
        setYomitanParserReadyPromise: () => {},
        getYomitanParserInitPromise: () => null,
        setYomitanParserInitPromise: () => {},
        isKnownWord: () => false,
        recordLookup: () => {},
        getKnownWordMatchMode: () => 'headword' as const,
        getNPlusOneEnabled: () => false,
        getMinSentenceWordsForNPlusOne: () => 3,
        getJlptLevel: () => null,
        getJlptEnabled: () => false,
        getFrequencyDictionaryEnabled: () => false,
        getFrequencyDictionaryMatchMode: () => 'headword' as const,
        getFrequencyRank: () => null,
        getYomitanGroupDebugEnabled: () => false,
        getMecabTokenizer: () => null,
      },
      createTokenizerRuntimeDeps: () => ({ isKnownWord: () => false }),
      tokenizeSubtitle: async (text: string) => ({ text }),
      createMecabTokenizerAndCheckMainDeps: {
        getMecabTokenizer: () => null,
        setMecabTokenizer: () => {},
        createMecabTokenizer: () => ({ id: 'mecab' }),
        checkAvailability: async () => {},
      },
      prewarmSubtitleDictionariesMainDeps: {
        ensureJlptDictionaryLookup: async () => {},
        ensureFrequencyDictionaryLookup: async () => {},
      },
    },
    warmups: {
      launchBackgroundWarmupTaskMainDeps: {
        now: () => 0,
        logDebug: () => {},
        logWarn: () => {},
      },
      startBackgroundWarmupsMainDeps: {
        getStarted: () => false,
        setStarted: () => {},
        isTexthookerOnlyMode: () => false,
        ensureYomitanExtensionLoaded: async () => {},
        shouldWarmupMecab: () => false,
        shouldWarmupYomitanExtension: () => false,
        shouldWarmupSubtitleDictionaries: () => false,
        shouldWarmupJellyfinRemoteSession: () => false,
        shouldAutoConnectJellyfinRemote: () => false,
        startJellyfinRemoteSession: async () => {},
      },
    },
  };
}

test('composeMpvRuntimeHandlers returns callable handlers and forwards to injected deps', async () => {
  const calls: string[] = [];
  let started = false;
  let metrics = BASE_METRICS;
  let mecabTokenizer: { id: string } | null = null;

  class FakeMpvClient {
    connected = false;

    constructor(
      public socketPath: string,
      public options: unknown,
    ) {
      const autoStartOverlay = (options as { autoStartOverlay: boolean }).autoStartOverlay;
      calls.push(`create-client:${socketPath}`);
      calls.push(`auto-start:${String(autoStartOverlay)}`);
    }

    on(): void {}

    connect(): void {
      this.connected = true;
      calls.push('client-connect');
    }
  }

  const fixture = createDefaultMpvFixture();
  const composed = composeMpvRuntimeHandlers<
    FakeMpvClient,
    { isKnownWord: (text: string) => boolean },
    { text: string }
  >({
    ...fixture,
    mpvClientRuntimeServiceFactoryMainDeps: {
      ...fixture.mpvClientRuntimeServiceFactoryMainDeps,
      createClient: FakeMpvClient,
      isAutoStartOverlayEnabled: () => true,
    },
    updateMpvSubtitleRenderMetricsMainDeps: {
      getCurrentMetrics: () => metrics,
      setCurrentMetrics: (next) => {
        metrics = next;
        calls.push('set-metrics');
      },
      applyPatch: (current, patch) => {
        calls.push('apply-metrics-patch');
        return { next: { ...current, ...patch }, changed: true };
      },
      broadcastMetrics: () => {
        calls.push('broadcast-metrics');
      },
    },
    tokenizer: {
      ...fixture.tokenizer,
      buildTokenizerDepsMainDeps: {
        ...fixture.tokenizer.buildTokenizerDepsMainDeps,
        isKnownWord: (text) => text === 'known',
        getJlptEnabled: () => true,
        getFrequencyDictionaryEnabled: () => true,
      },
      createTokenizerRuntimeDeps: (deps) => {
        calls.push('create-tokenizer-runtime-deps');
        return { isKnownWord: (text: string) => deps.isKnownWord(text) };
      },
      tokenizeSubtitle: async (text, deps) => {
        calls.push(`tokenize:${text}`);
        deps.isKnownWord('known');
        return { text };
      },
      createMecabTokenizerAndCheckMainDeps: {
        getMecabTokenizer: () => mecabTokenizer,
        setMecabTokenizer: (next) => {
          mecabTokenizer = next as { id: string };
          calls.push('set-mecab');
        },
        createMecabTokenizer: () => {
          calls.push('create-mecab');
          return { id: 'mecab' };
        },
        checkAvailability: async () => {
          calls.push('check-mecab');
        },
      },
      prewarmSubtitleDictionariesMainDeps: {
        ensureJlptDictionaryLookup: async () => {
          calls.push('prewarm-jlpt');
        },
        ensureFrequencyDictionaryLookup: async () => {
          calls.push('prewarm-frequency');
        },
      },
    },
    warmups: {
      launchBackgroundWarmupTaskMainDeps: {
        now: () => 100,
        logDebug: () => {
          calls.push('warmup-debug');
        },
        logWarn: () => {
          calls.push('warmup-warn');
        },
      },
      startBackgroundWarmupsMainDeps: {
        ...fixture.warmups.startBackgroundWarmupsMainDeps,
        getStarted: () => started,
        setStarted: (next) => {
          started = next;
          calls.push(`set-started:${String(next)}`);
        },
        ensureYomitanExtensionLoaded: async () => {
          calls.push('warmup-yomitan');
        },
        shouldWarmupMecab: () => true,
        shouldWarmupYomitanExtension: () => true,
        shouldWarmupSubtitleDictionaries: () => true,
        shouldWarmupJellyfinRemoteSession: () => true,
        startJellyfinRemoteSession: async () => {
          calls.push('warmup-jellyfin');
        },
      },
    },
  });

  assert.equal(typeof composed.bindMpvClientEventHandlers, 'function');
  assert.equal(typeof composed.createMpvClientRuntimeService, 'function');
  assert.equal(typeof composed.updateMpvSubtitleRenderMetrics, 'function');
  assert.equal(typeof composed.tokenizeSubtitle, 'function');
  assert.equal(typeof composed.createMecabTokenizerAndCheck, 'function');
  assert.equal(typeof composed.prewarmSubtitleDictionaries, 'function');
  assert.equal(typeof composed.startTokenizationWarmups, 'function');
  assert.equal(typeof composed.isTokenizationWarmupReady, 'function');
  assert.equal(typeof composed.launchBackgroundWarmupTask, 'function');
  assert.equal(typeof composed.startBackgroundWarmups, 'function');

  const client = composed.createMpvClientRuntimeService();
  assert.equal(client.connected, true);

  composed.updateMpvSubtitleRenderMetrics({ subPos: 90 });
  assert.equal(composed.isTokenizationWarmupReady(), false);
  await composed.startTokenizationWarmups();
  assert.equal(composed.isTokenizationWarmupReady(), true);
  const tokenized = await composed.tokenizeSubtitle('subtitle text');
  await composed.createMecabTokenizerAndCheck();
  await composed.prewarmSubtitleDictionaries();
  composed.startBackgroundWarmups();

  assert.deepEqual(tokenized, { text: 'subtitle text' });
  assert.equal(metrics.subPos, 90);
  assert.ok(calls.includes('create-client:/tmp/mpv.sock'));
  assert.ok(calls.includes('auto-start:true'));
  assert.ok(calls.includes('client-connect'));
  assert.ok(calls.includes('apply-metrics-patch'));
  assert.ok(calls.includes('set-metrics'));
  assert.ok(calls.includes('broadcast-metrics'));
  assert.ok(calls.includes('create-tokenizer-runtime-deps'));
  assert.ok(calls.includes('tokenize:subtitle text'));
  assert.ok(calls.includes('create-mecab'));
  assert.ok(calls.includes('set-mecab'));
  assert.ok(calls.includes('check-mecab'));
  assert.ok(calls.includes('prewarm-jlpt'));
  assert.ok(calls.includes('prewarm-frequency'));
  assert.ok(calls.includes('set-started:true'));
  assert.ok(calls.includes('warmup-yomitan'));
  assert.ok(calls.indexOf('create-mecab') < calls.indexOf('set-started:true'));
});

test('composeMpvRuntimeHandlers skips MeCab warmup when all POS-dependent annotations are disabled', async () => {
  const calls: string[] = [];
  let mecabTokenizer: { id: string } | null = null;

  class FakeMpvClient {
    connected = false;
    constructor(
      public socketPath: string,
      public options: unknown,
    ) {}
    on(): void {}
    connect(): void {
      this.connected = true;
    }
  }

  const fixture = createDefaultMpvFixture();
  const composed = composeMpvRuntimeHandlers<
    FakeMpvClient,
    { isKnownWord: (text: string) => boolean },
    { text: string }
  >({
    ...fixture,
    mpvClientRuntimeServiceFactoryMainDeps: {
      ...fixture.mpvClientRuntimeServiceFactoryMainDeps,
      createClient: FakeMpvClient,
      isAutoStartOverlayEnabled: () => true,
    },
    tokenizer: {
      ...fixture.tokenizer,
      createMecabTokenizerAndCheckMainDeps: {
        getMecabTokenizer: () => mecabTokenizer,
        setMecabTokenizer: (next) => {
          mecabTokenizer = next as { id: string };
          calls.push('set-mecab');
        },
        createMecabTokenizer: () => {
          calls.push('create-mecab');
          return { id: 'mecab' };
        },
        checkAvailability: async () => {
          calls.push('check-mecab');
        },
      },
    },
  });

  await composed.startTokenizationWarmups();

  assert.deepEqual(calls, []);
});

test('composeMpvRuntimeHandlers runs tokenization warmup once across sequential tokenize calls', async () => {
  let yomitanWarmupCalls = 0;
  let prewarmJlptCalls = 0;
  let prewarmFrequencyCalls = 0;
  const tokenizeCalls: string[] = [];

  const fixture = createDefaultMpvFixture();
  const composed = composeMpvRuntimeHandlers<
    { connect: () => void; on: () => void },
    { isKnownWord: () => boolean },
    { text: string }
  >({
    ...fixture,
    tokenizer: {
      ...fixture.tokenizer,
      tokenizeSubtitle: async (text) => {
        tokenizeCalls.push(text);
        return { text };
      },
      prewarmSubtitleDictionariesMainDeps: {
        ensureJlptDictionaryLookup: async () => {
          prewarmJlptCalls += 1;
        },
        ensureFrequencyDictionaryLookup: async () => {
          prewarmFrequencyCalls += 1;
        },
      },
    },
    warmups: {
      ...fixture.warmups,
      startBackgroundWarmupsMainDeps: {
        ...fixture.warmups.startBackgroundWarmupsMainDeps,
        ensureYomitanExtensionLoaded: async () => {
          yomitanWarmupCalls += 1;
        },
      },
    },
  });

  await composed.tokenizeSubtitle('first');
  await composed.tokenizeSubtitle('second');

  assert.deepEqual(tokenizeCalls, ['first', 'second']);
  assert.equal(yomitanWarmupCalls, 1);
  assert.equal(prewarmJlptCalls, 0);
  assert.equal(prewarmFrequencyCalls, 0);
});

test('composeMpvRuntimeHandlers does not block first tokenization on dictionary or MeCab warmup', async () => {
  const jlptDeferred = createDeferred();
  const frequencyDeferred = createDeferred();
  const mecabDeferred = createDeferred();
  let tokenizeResolved = false;

  const fixture = createDefaultMpvFixture();
  const composed = composeMpvRuntimeHandlers<
    { connect: () => void; on: () => void },
    { isKnownWord: () => boolean },
    { text: string }
  >({
    ...fixture,
    tokenizer: {
      ...fixture.tokenizer,
      buildTokenizerDepsMainDeps: {
        ...fixture.tokenizer.buildTokenizerDepsMainDeps,
        getNPlusOneEnabled: () => true,
        getJlptEnabled: () => true,
        getFrequencyDictionaryEnabled: () => true,
      },
      createMecabTokenizerAndCheckMainDeps: {
        ...fixture.tokenizer.createMecabTokenizerAndCheckMainDeps,
        checkAvailability: async () => mecabDeferred.promise,
      },
      prewarmSubtitleDictionariesMainDeps: {
        ensureJlptDictionaryLookup: async () => jlptDeferred.promise,
        ensureFrequencyDictionaryLookup: async () => frequencyDeferred.promise,
      },
    },
  });

  const tokenizePromise = composed.tokenizeSubtitle('first line').then(() => {
    tokenizeResolved = true;
  });
  await new Promise<void>((resolve) => setImmediate(resolve));
  assert.equal(tokenizeResolved, true);

  jlptDeferred.resolve();
  frequencyDeferred.resolve();
  mecabDeferred.resolve();
  await tokenizePromise;
  await composed.startTokenizationWarmups();
});

test('composeMpvRuntimeHandlers shows annotation loading OSD after tokenization-ready when dictionary warmup is still pending', async () => {
  const jlptDeferred = createDeferred();
  const frequencyDeferred = createDeferred();
  const osdMessages: string[] = [];

  const fixture = createDefaultMpvFixture();
  const composed = composeMpvRuntimeHandlers<
    { connect: () => void; on: () => void },
    { onTokenizationReady?: (text: string) => void },
    { text: string }
  >({
    ...fixture,
    tokenizer: {
      ...fixture.tokenizer,
      buildTokenizerDepsMainDeps: {
        ...fixture.tokenizer.buildTokenizerDepsMainDeps,
        getJlptEnabled: () => true,
        getFrequencyDictionaryEnabled: () => true,
      },
      createTokenizerRuntimeDeps: (deps) =>
        deps as unknown as { onTokenizationReady?: (text: string) => void },
      tokenizeSubtitle: async (text, deps) => {
        deps.onTokenizationReady?.(text);
        return { text };
      },
      prewarmSubtitleDictionariesMainDeps: {
        ensureJlptDictionaryLookup: async () => jlptDeferred.promise,
        ensureFrequencyDictionaryLookup: async () => frequencyDeferred.promise,
        showMpvOsd: (message) => {
          osdMessages.push(message);
        },
      },
    },
  });

  const warmupPromise = composed.startTokenizationWarmups();
  await new Promise<void>((resolve) => setImmediate(resolve));
  assert.deepEqual(osdMessages, []);
  assert.equal(composed.isTokenizationWarmupReady(), false);

  await composed.tokenizeSubtitle('first line');
  assert.deepEqual(osdMessages, ['Loading subtitle annotations |']);
  assert.equal(composed.isTokenizationWarmupReady(), true);

  jlptDeferred.resolve();
  frequencyDeferred.resolve();
  await warmupPromise;
  await new Promise<void>((resolve) => setImmediate(resolve));

  assert.deepEqual(osdMessages, ['Loading subtitle annotations |', 'Subtitle annotations loaded']);
});

test('composeMpvRuntimeHandlers reuses completed background tokenization warmups for later tokenize calls', async () => {
  let started = false;
  let yomitanWarmupCalls = 0;
  let mecabWarmupCalls = 0;
  let jlptWarmupCalls = 0;
  let frequencyWarmupCalls = 0;
  let mecabTokenizer: { tokenize: () => Promise<never[]> } | null = null;

  const fixture = createDefaultMpvFixture();
  const composed = composeMpvRuntimeHandlers<
    { connect: () => void; on: () => void },
    { isKnownWord: () => boolean },
    { text: string }
  >({
    ...fixture,
    tokenizer: {
      ...fixture.tokenizer,
      buildTokenizerDepsMainDeps: {
        ...fixture.tokenizer.buildTokenizerDepsMainDeps,
        getNPlusOneEnabled: () => true,
        getJlptEnabled: () => true,
        getFrequencyDictionaryEnabled: () => true,
        getMecabTokenizer: () => mecabTokenizer,
      },
      createMecabTokenizerAndCheckMainDeps: {
        getMecabTokenizer: () => mecabTokenizer,
        setMecabTokenizer: (next) => {
          mecabTokenizer = next as { tokenize: () => Promise<never[]> };
        },
        createMecabTokenizer: () => ({ tokenize: async () => [] }),
        checkAvailability: async () => {
          mecabWarmupCalls += 1;
        },
      },
      prewarmSubtitleDictionariesMainDeps: {
        ensureJlptDictionaryLookup: async () => {
          jlptWarmupCalls += 1;
        },
        ensureFrequencyDictionaryLookup: async () => {
          frequencyWarmupCalls += 1;
        },
      },
    },
    warmups: {
      ...fixture.warmups,
      startBackgroundWarmupsMainDeps: {
        ...fixture.warmups.startBackgroundWarmupsMainDeps,
        getStarted: () => started,
        setStarted: (next) => {
          started = next;
        },
        ensureYomitanExtensionLoaded: async () => {
          yomitanWarmupCalls += 1;
        },
        shouldWarmupMecab: () => true,
        shouldWarmupYomitanExtension: () => true,
        shouldWarmupSubtitleDictionaries: () => true,
      },
    },
  });

  composed.startBackgroundWarmups();
  await new Promise<void>((resolve) => setImmediate(resolve));

  assert.equal(yomitanWarmupCalls, 1);
  assert.equal(mecabWarmupCalls, 1);
  assert.equal(jlptWarmupCalls, 1);
  assert.equal(frequencyWarmupCalls, 1);

  await composed.tokenizeSubtitle('first line after background warmup');

  assert.equal(yomitanWarmupCalls, 1);
  assert.equal(mecabWarmupCalls, 1);
  assert.equal(jlptWarmupCalls, 1);
  assert.equal(frequencyWarmupCalls, 1);
});
