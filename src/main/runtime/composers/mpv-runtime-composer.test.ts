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

  const composed = composeMpvRuntimeHandlers<
    FakeMpvClient,
    { isKnownWord: (text: string) => boolean },
    { text: string }
  >({
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
      createClient: FakeMpvClient,
      getSocketPath: () => '/tmp/mpv.sock',
      getResolvedConfig: () => ({ auto_start_overlay: false }),
      isAutoStartOverlayEnabled: () => true,
      setOverlayVisible: () => {},
      isVisibleOverlayVisible: () => false,
      getReconnectTimer: () => null,
      setReconnectTimer: () => {},
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
      buildTokenizerDepsMainDeps: {
        getYomitanExt: () => null,
        getYomitanParserWindow: () => null,
        setYomitanParserWindow: () => {},
        getYomitanParserReadyPromise: () => null,
        setYomitanParserReadyPromise: () => {},
        getYomitanParserInitPromise: () => null,
        setYomitanParserInitPromise: () => {},
        isKnownWord: (text) => text === 'known',
        recordLookup: () => {},
        getKnownWordMatchMode: () => 'headword',
        getMinSentenceWordsForNPlusOne: () => 3,
        getJlptLevel: () => null,
        getJlptEnabled: () => true,
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: () => null,
        getYomitanGroupDebugEnabled: () => false,
        getMecabTokenizer: () => null,
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
        getStarted: () => started,
        setStarted: (next) => {
          started = next;
          calls.push(`set-started:${String(next)}`);
        },
        isTexthookerOnlyMode: () => false,
        ensureYomitanExtensionLoaded: async () => {
          calls.push('warmup-yomitan');
        },
        shouldWarmupMecab: () => true,
        shouldWarmupYomitanExtension: () => true,
        shouldWarmupSubtitleDictionaries: () => true,
        shouldWarmupJellyfinRemoteSession: () => true,
        shouldAutoConnectJellyfinRemote: () => false,
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
  assert.equal(typeof composed.launchBackgroundWarmupTask, 'function');
  assert.equal(typeof composed.startBackgroundWarmups, 'function');

  const client = composed.createMpvClientRuntimeService();
  assert.equal(client.connected, true);

  composed.updateMpvSubtitleRenderMetrics({ subPos: 90 });
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
