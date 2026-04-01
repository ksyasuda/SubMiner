import * as fs from 'fs';
import * as path from 'path';
import { MecabTokenizer } from '../mecab-tokenizer';
import {
  MpvIpcClient,
  applyMpvSubtitleRenderMetricsPatch,
  createShiftSubtitleDelayToAdjacentCueHandler,
  createTokenizerDepsRuntime,
  cycleSecondarySubMode as cycleSecondarySubModeCore,
  sendMpvCommandRuntime,
  showMpvOsdRuntime,
  tokenizeSubtitle as tokenizeSubtitleCore,
} from '../core/services';
import type {
  MpvSubtitleRenderMetrics,
  ResolvedConfig,
  SecondarySubMode,
  SubtitleData,
} from '../types';
import type { AppState } from './state';
import type { SubtitleRuntime } from './subtitle-runtime';
import type { createCurrentMediaTokenizationGate } from './runtime/current-media-tokenization-gate';
import type { createStartupOsdSequencer } from './runtime/startup-osd-sequencer';
import { createMpvOsdRuntimeHandlers } from './runtime/mpv-osd-runtime-handlers';
import { createCycleSecondarySubModeRuntimeHandler } from './runtime/secondary-sub-mode-runtime-handler';
import { composeMpvRuntimeHandlers } from './runtime/composers';

type RuntimeOptionId =
  | 'subtitle.annotation.nPlusOne'
  | 'subtitle.annotation.jlpt'
  | 'subtitle.annotation.frequency';

interface MpvRuntimeLogger {
  debug: (message: string, meta?: unknown) => void;
  info: (message: string, meta?: unknown) => void;
  warn: (message: string, meta?: unknown) => void;
  error: (message: string, error?: unknown) => void;
}

export interface MpvRuntimeInput {
  appState: AppState;
  logPath: string;
  logger: MpvRuntimeLogger;
  getResolvedConfig: () => ResolvedConfig;
  getRuntimeBooleanOption: (id: RuntimeOptionId, fallback: boolean) => boolean;
  subtitle: Pick<
    SubtitleRuntime,
    | 'consumeCachedSubtitle'
    | 'emitSubtitlePayload'
    | 'onSubtitleChange'
    | 'onCurrentMediaPathChange'
    | 'onTimePosUpdate'
    | 'scheduleSubtitlePrefetchRefresh'
    | 'loadSubtitleSourceText'
    | 'setTokenizeSubtitleDeferred'
  >;
  ensureYomitanExtensionLoaded: () => Promise<void>;
  currentMediaTokenizationGate: ReturnType<typeof createCurrentMediaTokenizationGate>;
  startupOsdSequencer: ReturnType<typeof createStartupOsdSequencer>;
  dictionaries: {
    ensureJlptDictionaryLookup: () => Promise<void>;
    ensureFrequencyDictionaryLookup: () => Promise<void>;
  };
  mediaRuntime: {
    syncImmersionMediaState: () => void;
    updateCurrentMediaPath: (mediaPath: unknown) => void;
    updateCurrentMediaTitle: (mediaTitle: unknown) => void;
  };
  characterDictionaryAutoSyncRuntime: {
    scheduleSync: () => void;
  };
  overlay: {
    broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
    getVisibleOverlayVisible: () => boolean;
    setOverlayVisible: (visible: boolean) => void;
  };
  lifecycle: {
    requestAppQuit: () => void;
    scheduleQuitCheck: (callback: () => void) => void;
    restoreOverlayMpvSubtitles: () => void;
    syncOverlayMpvSubtitleSuppression: () => void;
    refreshDiscordPresence: () => void;
  };
  stats: {
    ensureImmersionTrackerStarted: () => void;
  };
  anilist: {
    getCurrentAnilistMediaKey: () => string | null;
    resetAnilistMediaTracking: (mediaKey: string | null) => void;
    maybeProbeAnilistDuration: (mediaKey: string | null) => void;
    ensureAnilistMediaGuess: (mediaKey: string | null) => void;
    maybeRunAnilistPostWatchUpdate: () => Promise<void>;
    resetAnilistMediaGuessState: () => void;
  };
  jellyfin: {
    getQuitOnDisconnectArmed: () => boolean;
    reportJellyfinRemoteStopped: () => Promise<void>;
    reportJellyfinRemoteProgress: (forceImmediate?: boolean) => Promise<void>;
    startJellyfinRemoteSession: () => Promise<void>;
  };
  youtube: {
    getQuitOnDisconnectArmed: () => boolean;
    handleMpvConnectionChange: (connected: boolean) => void;
    handleMediaPathChange: (path: string | null) => void;
    handleSubtitleTrackChange: (sid: number | null) => void;
    handleSubtitleTrackListChange: (trackList: unknown[] | null) => void;
    invalidatePendingAutoplayReadyFallbacks: () => void;
    maybeSignalPluginAutoplayReady: (
      subtitle: { text: string; tokens: null },
      options?: { forceWhilePaused?: boolean },
    ) => void;
  };
  isCharacterDictionaryEnabled: () => boolean;
}

export interface MpvRuntime {
  createMpvClientRuntimeService: () => MpvIpcClient;
  updateMpvSubtitleRenderMetrics: (patch: Partial<MpvSubtitleRenderMetrics>) => void;
  createMecabTokenizerAndCheck: () => Promise<void>;
  prewarmSubtitleDictionaries: () => Promise<void>;
  startTokenizationWarmups: () => Promise<void>;
  isTokenizationWarmupReady: () => boolean;
  startBackgroundWarmups: () => void;
  showMpvOsd: (text: string) => void;
  flushMpvLog: () => Promise<void>;
  cycleSecondarySubMode: () => void;
  shiftSubtitleDelayToAdjacentCue: (direction: 'next' | 'previous') => Promise<void>;
}

function getActiveMediaPath(appState: AppState): string | null {
  return appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null;
}

export function createMpvRuntime(input: MpvRuntimeInput): MpvRuntime {
  let backgroundWarmupsStarted = false;
  let tokenizeSubtitleDeferred: ((text: string) => Promise<SubtitleData>) | null = null;

  const { flushMpvLog, showMpvOsd } = createMpvOsdRuntimeHandlers({
    appendToMpvLogMainDeps: {
      logPath: input.logPath,
      dirname: (targetPath) => path.dirname(targetPath),
      mkdir: async (targetPath, options) => {
        await fs.promises.mkdir(targetPath, options);
      },
      appendFile: async (targetPath, data, options) => {
        await fs.promises.appendFile(targetPath, data, options);
      },
      now: () => new Date(),
    },
    buildShowMpvOsdMainDeps: (appendToMpvLog) => ({
      appendToMpvLog,
      showMpvOsdRuntime: (mpvClient, text, fallbackLog) =>
        showMpvOsdRuntime(mpvClient, text, fallbackLog),
      getMpvClient: () => input.appState.mpvClient,
      logInfo: (line) => input.logger.info(line),
    }),
  });

  const cycleSecondarySubMode = createCycleSecondarySubModeRuntimeHandler({
    cycleSecondarySubModeMainDeps: {
      getSecondarySubMode: () => input.appState.secondarySubMode,
      setSecondarySubMode: (mode: SecondarySubMode) => {
        input.appState.secondarySubMode = mode;
      },
      getLastSecondarySubToggleAtMs: () => input.appState.lastSecondarySubToggleAtMs,
      setLastSecondarySubToggleAtMs: (timestampMs: number) => {
        input.appState.lastSecondarySubToggleAtMs = timestampMs;
      },
      broadcastToOverlayWindows: (channel, mode) => {
        input.overlay.broadcastToOverlayWindows(channel, mode);
      },
      showMpvOsd: (text: string) => showMpvOsd(text),
    },
    cycleSecondarySubMode: (deps) => cycleSecondarySubModeCore(deps),
  });

  const {
    createMpvClientRuntimeService: createMpvClientRuntimeServiceHandler,
    updateMpvSubtitleRenderMetrics: updateMpvSubtitleRenderMetricsHandler,
    tokenizeSubtitle,
    createMecabTokenizerAndCheck,
    prewarmSubtitleDictionaries,
    startBackgroundWarmups,
    startTokenizationWarmups,
    isTokenizationWarmupReady,
  } = composeMpvRuntimeHandlers<
    MpvIpcClient,
    ReturnType<typeof createTokenizerDepsRuntime>,
    SubtitleData
  >({
    bindMpvMainEventHandlersMainDeps: {
      appState: input.appState,
      getQuitOnDisconnectArmed: () =>
        input.jellyfin.getQuitOnDisconnectArmed() || input.youtube.getQuitOnDisconnectArmed(),
      scheduleQuitCheck: (callback) => {
        input.lifecycle.scheduleQuitCheck(callback);
      },
      quitApp: () => input.lifecycle.requestAppQuit(),
      reportJellyfinRemoteStopped: () => {
        void input.jellyfin.reportJellyfinRemoteStopped();
      },
      maybeRunAnilistPostWatchUpdate: () => input.anilist.maybeRunAnilistPostWatchUpdate(),
      logSubtitleTimingError: (message, error) => input.logger.error(message, error),
      broadcastToOverlayWindows: (channel, payload) => {
        input.overlay.broadcastToOverlayWindows(channel, payload);
      },
      getImmediateSubtitlePayload: (text) => input.subtitle.consumeCachedSubtitle(text),
      emitImmediateSubtitle: (payload) => {
        input.subtitle.emitSubtitlePayload(payload);
      },
      onSubtitleChange: (text) => {
        input.subtitle.onSubtitleChange(text);
      },
      refreshDiscordPresence: () => {
        input.lifecycle.refreshDiscordPresence();
      },
      ensureImmersionTrackerInitialized: () => {
        input.stats.ensureImmersionTrackerStarted();
      },
      tokenizeSubtitleForImmersion: async (text): Promise<SubtitleData | null> =>
        tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
      updateCurrentMediaPath: (mediaPath) => {
        input.youtube.invalidatePendingAutoplayReadyFallbacks();
        input.currentMediaTokenizationGate.updateCurrentMediaPath(mediaPath);
        input.startupOsdSequencer.reset();
        input.subtitle.onCurrentMediaPathChange(mediaPath);
        input.youtube.handleMediaPathChange(mediaPath);
        if (mediaPath) {
          input.stats.ensureImmersionTrackerStarted();
        }
        input.mediaRuntime.updateCurrentMediaPath(mediaPath);
      },
      restoreMpvSubVisibility: () => {
        input.lifecycle.restoreOverlayMpvSubtitles();
      },
      resetSubtitleSidebarEmbeddedLayout: () => {
        sendMpvCommandRuntime(input.appState.mpvClient, [
          'set_property',
          'video-margin-ratio-right',
          0,
        ]);
        sendMpvCommandRuntime(input.appState.mpvClient, ['set_property', 'video-pan-x', 0]);
      },
      getCurrentAnilistMediaKey: () => input.anilist.getCurrentAnilistMediaKey(),
      resetAnilistMediaTracking: (mediaKey) => {
        input.anilist.resetAnilistMediaTracking(mediaKey);
      },
      maybeProbeAnilistDuration: (mediaKey) => {
        if (mediaKey) {
          input.anilist.maybeProbeAnilistDuration(mediaKey);
        }
      },
      ensureAnilistMediaGuess: (mediaKey) => {
        if (mediaKey) {
          input.anilist.ensureAnilistMediaGuess(mediaKey);
        }
      },
      syncImmersionMediaState: () => {
        input.mediaRuntime.syncImmersionMediaState();
      },
      signalAutoplayReadyIfWarm: () => {
        if (!isTokenizationWarmupReady()) {
          return;
        }
        input.youtube.maybeSignalPluginAutoplayReady(
          { text: '__warm__', tokens: null },
          { forceWhilePaused: true },
        );
      },
      scheduleCharacterDictionarySync: () => {
        if (!input.isCharacterDictionaryEnabled()) {
          return;
        }
        input.characterDictionaryAutoSyncRuntime.scheduleSync();
      },
      updateCurrentMediaTitle: (title) => {
        input.mediaRuntime.updateCurrentMediaTitle(title);
      },
      resetAnilistMediaGuessState: () => {
        input.anilist.resetAnilistMediaGuessState();
      },
      reportJellyfinRemoteProgress: (forceImmediate) => {
        void input.jellyfin.reportJellyfinRemoteProgress(forceImmediate);
      },
      onTimePosUpdate: (time) => {
        input.subtitle.onTimePosUpdate(time);
      },
      onSubtitleTrackChange: (sid) => {
        input.subtitle.scheduleSubtitlePrefetchRefresh();
        input.youtube.handleSubtitleTrackChange(sid);
      },
      onSubtitleTrackListChange: (trackList) => {
        input.subtitle.scheduleSubtitlePrefetchRefresh();
        input.youtube.handleSubtitleTrackListChange(trackList);
      },
      updateSubtitleRenderMetrics: (patch) => {
        updateMpvSubtitleRenderMetricsHandler(patch as Partial<MpvSubtitleRenderMetrics>);
      },
      syncOverlayMpvSubtitleSuppression: () => {
        input.lifecycle.syncOverlayMpvSubtitleSuppression();
      },
    },
    mpvClientRuntimeServiceFactoryMainDeps: {
      createClient: MpvIpcClient,
      getSocketPath: () => input.appState.mpvSocketPath,
      getResolvedConfig: () => input.getResolvedConfig(),
      isAutoStartOverlayEnabled: () => input.appState.autoStartOverlay,
      setOverlayVisible: (visible: boolean) => {
        input.overlay.setOverlayVisible(visible);
      },
      isVisibleOverlayVisible: () => input.overlay.getVisibleOverlayVisible(),
      getReconnectTimer: () => input.appState.reconnectTimer,
      setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => {
        input.appState.reconnectTimer = timer;
      },
    },
    updateMpvSubtitleRenderMetricsMainDeps: {
      getCurrentMetrics: () => input.appState.mpvSubtitleRenderMetrics,
      setCurrentMetrics: (metrics) => {
        input.appState.mpvSubtitleRenderMetrics = metrics;
      },
      applyPatch: (current, patch) => applyMpvSubtitleRenderMetricsPatch(current, patch),
      broadcastMetrics: () => {},
    },
    tokenizer: {
      buildTokenizerDepsMainDeps: {
        getYomitanExt: () => input.appState.yomitanExt,
        getYomitanSession: () => input.appState.yomitanSession,
        getYomitanParserWindow: () => input.appState.yomitanParserWindow,
        setYomitanParserWindow: (window) => {
          input.appState.yomitanParserWindow = window;
        },
        getYomitanParserReadyPromise: () => input.appState.yomitanParserReadyPromise,
        setYomitanParserReadyPromise: (promise) => {
          input.appState.yomitanParserReadyPromise = promise;
        },
        getYomitanParserInitPromise: () => input.appState.yomitanParserInitPromise,
        setYomitanParserInitPromise: (promise) => {
          input.appState.yomitanParserInitPromise = promise;
        },
        isKnownWord: (text) => Boolean(input.appState.ankiIntegration?.isKnownWord(text)),
        recordLookup: (hit) => {
          input.stats.ensureImmersionTrackerStarted();
          input.appState.immersionTracker?.recordLookup(hit);
        },
        getKnownWordMatchMode: () =>
          input.appState.ankiIntegration?.getKnownWordMatchMode() ??
          input.getResolvedConfig().ankiConnect.knownWords.matchMode,
        getNPlusOneEnabled: () =>
          input.getRuntimeBooleanOption(
            'subtitle.annotation.nPlusOne',
            input.getResolvedConfig().ankiConnect.knownWords.highlightEnabled,
          ),
        getMinSentenceWordsForNPlusOne: () =>
          input.getResolvedConfig().ankiConnect.nPlusOne.minSentenceWords,
        getJlptLevel: (text) => input.appState.jlptLevelLookup(text),
        getJlptEnabled: () =>
          input.getRuntimeBooleanOption(
            'subtitle.annotation.jlpt',
            input.getResolvedConfig().subtitleStyle.enableJlpt,
          ),
        getCharacterDictionaryEnabled: () => input.isCharacterDictionaryEnabled(),
        getNameMatchEnabled: () => input.getResolvedConfig().subtitleStyle.nameMatchEnabled,
        getFrequencyDictionaryEnabled: () =>
          input.getRuntimeBooleanOption(
            'subtitle.annotation.frequency',
            input.getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
          ),
        getFrequencyDictionaryMatchMode: () =>
          input.getResolvedConfig().subtitleStyle.frequencyDictionary.matchMode,
        getFrequencyRank: (text) => input.appState.frequencyRankLookup(text),
        getYomitanGroupDebugEnabled: () => input.appState.overlayDebugVisualizationEnabled,
        getMecabTokenizer: () => input.appState.mecabTokenizer,
        onTokenizationReady: (text) => {
          input.currentMediaTokenizationGate.markReady(getActiveMediaPath(input.appState));
          input.startupOsdSequencer.markTokenizationReady();
          input.youtube.maybeSignalPluginAutoplayReady(
            { text, tokens: null },
            { forceWhilePaused: true },
          );
        },
      },
      createTokenizerRuntimeDeps: (deps) =>
        createTokenizerDepsRuntime(deps as Parameters<typeof createTokenizerDepsRuntime>[0]),
      tokenizeSubtitle: (text, deps) => tokenizeSubtitleCore(text, deps),
      createMecabTokenizerAndCheckMainDeps: {
        getMecabTokenizer: () => input.appState.mecabTokenizer,
        setMecabTokenizer: (tokenizer) => {
          input.appState.mecabTokenizer = tokenizer as MecabTokenizer | null;
        },
        createMecabTokenizer: () => new MecabTokenizer(),
        checkAvailability: async (tokenizer) => (tokenizer as MecabTokenizer).checkAvailability(),
      },
      prewarmSubtitleDictionariesMainDeps: {
        ensureJlptDictionaryLookup: () => input.dictionaries.ensureJlptDictionaryLookup(),
        ensureFrequencyDictionaryLookup: () => input.dictionaries.ensureFrequencyDictionaryLookup(),
        showMpvOsd: (message: string) => showMpvOsd(message),
        showLoadingOsd: (message: string) =>
          input.startupOsdSequencer.showAnnotationLoading(message),
        showLoadedOsd: (message: string) =>
          input.startupOsdSequencer.markAnnotationLoadingComplete(message),
        shouldShowOsdNotification: () => {
          const type = input.getResolvedConfig().ankiConnect.behavior.notificationType;
          return type === 'osd' || type === 'both';
        },
      },
    },
    warmups: {
      launchBackgroundWarmupTaskMainDeps: {
        now: () => Date.now(),
        logDebug: (message) => input.logger.debug(message),
        logWarn: (message) => input.logger.warn(message),
      },
      startBackgroundWarmupsMainDeps: {
        getStarted: () => backgroundWarmupsStarted,
        setStarted: (started) => {
          backgroundWarmupsStarted = started;
        },
        isTexthookerOnlyMode: () => input.appState.texthookerOnlyMode,
        ensureYomitanExtensionLoaded: () => input.ensureYomitanExtensionLoaded().then(() => {}),
        shouldWarmupMecab: () => {
          const startupWarmups = input.getResolvedConfig().startupWarmups;
          if (startupWarmups.lowPowerMode) {
            return false;
          }
          if (!startupWarmups.mecab) {
            return false;
          }
          return (
            input.getRuntimeBooleanOption(
              'subtitle.annotation.nPlusOne',
              input.getResolvedConfig().ankiConnect.knownWords.highlightEnabled,
            ) ||
            input.getRuntimeBooleanOption(
              'subtitle.annotation.jlpt',
              input.getResolvedConfig().subtitleStyle.enableJlpt,
            ) ||
            input.getRuntimeBooleanOption(
              'subtitle.annotation.frequency',
              input.getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
            )
          );
        },
        shouldWarmupYomitanExtension: () =>
          input.getResolvedConfig().startupWarmups.yomitanExtension,
        shouldWarmupSubtitleDictionaries: () => {
          const startupWarmups = input.getResolvedConfig().startupWarmups;
          if (startupWarmups.lowPowerMode) {
            return false;
          }
          return startupWarmups.subtitleDictionaries;
        },
        shouldWarmupJellyfinRemoteSession: () => {
          const startupWarmups = input.getResolvedConfig().startupWarmups;
          if (startupWarmups.lowPowerMode) {
            return false;
          }
          return startupWarmups.jellyfinRemoteSession;
        },
        shouldAutoConnectJellyfinRemote: () => {
          const jellyfin = input.getResolvedConfig().jellyfin;
          return (
            jellyfin.enabled && jellyfin.remoteControlEnabled && jellyfin.remoteControlAutoConnect
          );
        },
        startJellyfinRemoteSession: () => input.jellyfin.startJellyfinRemoteSession(),
        logDebug: (message) => input.logger.debug(message),
      },
    },
  });
  tokenizeSubtitleDeferred = tokenizeSubtitle;
  input.subtitle.setTokenizeSubtitleDeferred(tokenizeSubtitle);

  const createMpvClientRuntimeService = (): MpvIpcClient => {
    const client = createMpvClientRuntimeServiceHandler();
    client.on('connection-change', ({ connected }) => {
      input.youtube.handleMpvConnectionChange(connected);
    });
    return client;
  };

  const shiftSubtitleDelayToAdjacentCue = createShiftSubtitleDelayToAdjacentCueHandler({
    getMpvClient: () => input.appState.mpvClient,
    loadSubtitleSourceText: (source) => input.subtitle.loadSubtitleSourceText(source),
    sendMpvCommand: (command) => sendMpvCommandRuntime(input.appState.mpvClient, command),
    showMpvOsd: (text) => showMpvOsd(text),
  });

  return {
    createMpvClientRuntimeService,
    updateMpvSubtitleRenderMetrics: (patch) => {
      updateMpvSubtitleRenderMetricsHandler(patch);
    },
    createMecabTokenizerAndCheck,
    prewarmSubtitleDictionaries,
    startTokenizationWarmups,
    isTokenizationWarmupReady,
    startBackgroundWarmups,
    showMpvOsd,
    flushMpvLog,
    cycleSecondarySubMode,
    shiftSubtitleDelayToAdjacentCue,
  };
}
