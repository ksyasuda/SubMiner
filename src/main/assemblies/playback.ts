import {
  MpvIpcClient,
  applyMpvSubtitleRenderMetricsPatch,
  createTokenizerDepsRuntime,
  isSubtitleAnnotationUpgrade,
  resolveSanitizedSubtitleSeekCommand,
  sendMpvCommandRuntime,
  toSubtitleClock,
  tokenizeSubtitle as tokenizeSubtitleCore,
} from '../../core/services';
import { parseSubtitleCues } from '../../core/services/subtitle-cue-parser';
import { resolveAniSkipMetadataForFile } from '../runtime/aniskip-metadata';
import { createAniSkipRuntime } from '../runtime/aniskip-runtime';
import { resolveCurrentSubtitleForRenderer } from '../runtime/current-subtitle-snapshot';
import { resolveFreshPlaybackPaused } from '../runtime/playback-paused-state';
import { createAutoplaySubtitlePrimingRuntime } from '../runtime/autoplay-subtitle-priming-runtime';
import { createAutoplayTokenizationWarmRelease } from '../runtime/autoplay-tokenization-warm-release';
import { composeMpvRuntimeHandlers } from '../runtime/composers';
import { createCachedInternalSubtitleTrackExtractor } from '../runtime/internal-subtitle-extraction';
import {
  markJellyfinRemotePlaybackLoaded as markJellyfinRemotePlaybackLoadedState,
  shouldAutoLoadSecondarySubTrackForJellyfinPlayback,
} from '../runtime/jellyfin-remote-playback';
import { buildSubtitleTrackDiagnostics } from '../runtime/mpv-track-diagnostics';
import { createRemoteMediaPathDetector } from '../runtime/network-media-path';
import { resolveCanonicalPrimarySubtitle } from '../runtime/primary-subtitle-text';
import { createSecondarySubtitleTrackController } from '../runtime/secondary-subtitle-track';
import { shouldQuitOnMpvShutdownForTrayState } from '../runtime/startup-tray-policy';
import { createSubtitlePrefetchInitController } from '../runtime/subtitle-prefetch-init';
import {
  createRefreshSubtitlePrefetchFromActiveTrackHandler,
  createResolveActiveSubtitleSidebarSourceHandler,
} from '../runtime/subtitle-prefetch-runtime';
import { createPrefetchAwareSubtitleProcessingController } from '../runtime/subtitle-processing-runtime';
import {
  isSameYoutubeMediaPath,
  isYoutubeMediaPath,
  shouldUseCachedYoutubeParsedCues,
} from '../runtime/youtube-playback';
import { MecabTokenizer } from '../../mecab-tokenizer';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type {
  MpvSubtitleRenderMetrics,
  SubtitleCue,
  SubtitleData,
  SubtitleSidebarPlaybackState,
} from '../../types';
import type { BrowserWindow } from 'electron';
import type { ConfigService } from '../../config/service';
import type { createOverlayManager, SubtitleWebSocket } from '../../core/services';
import type {
  SubtitlePrefetchService,
  SubtitlePrefetchServiceDeps,
} from '../../core/services/subtitle-prefetch';
import type { Logger } from '../../logger';
import type { createMediaRuntimeService } from '../media-runtime';
import type { createAutoplayReadyGate } from '../runtime/autoplay-ready-gate';
import type { createCurrentMediaTokenizationGate } from '../runtime/current-media-tokenization-gate';
import type { createDiscordPresenceRuntime } from '../runtime/discord-presence-runtime';
import type { createImmersionMediaRuntime } from '../runtime/domains/startup';
import type { loadSubtitleSourceText } from '../runtime/internal-subtitle-extraction';
import type { createManagedLocalSubtitleSelectionRuntime } from '../runtime/local-subtitle-selection';
import type { createMediaTimingReviewRuntime } from '../runtime/media-timing-review';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import type { createStartupOsdSequencer } from '../runtime/startup-osd-sequencer';
import type { createSubtitleGenerationRuntime } from '../runtime/subtitle-generation-runtime';
import type { RuntimeOptionId } from '../../types/runtime-options';
import type { OverlayModalRuntime } from '../overlay-runtime';
import type { AppState } from '../state';
import type { AnilistAssembly } from './anilist';
import type { CharacterDictionaryAssembly } from './character-dictionary';
import type { DictionariesAssembly } from './dictionaries';
import type { JellyfinAssembly } from './jellyfin';
import type { IpcSlice } from './slices';
import type { YoutubeAssembly } from './youtube';

type SubtitleWsBroadcaster = Pick<SubtitleWebSocket, 'broadcast'>;
type OverlayNotificationsRuntime = ReturnType<typeof createOverlayNotificationsRuntime>;
type AnnotationRuntimeOptionId = Extract<
  RuntimeOptionId,
  | 'subtitle.annotation.knownWords.highlightEnabled'
  | 'subtitle.annotation.nPlusOne'
  | 'subtitle.annotation.jlpt'
  | 'subtitle.annotation.frequency'
>;

/**
 * What the playback assembly needs from the rest of `main.ts`. Everything here except the
 * getters is built before the assembly; siblings built later are reached through getters, and
 * construction calls none of them.
 */
export interface PlaybackAssemblyDeps {
  appState: AppState;
  configService: Pick<ConfigService, 'getConfig'>;
  logger: Pick<Logger, 'debug' | 'info' | 'warn' | 'error'>;
  overlayManager: Pick<
    ReturnType<typeof createOverlayManager>,
    'broadcastToOverlayWindows' | 'getVisibleOverlayVisible'
  >;
  subtitleWsService: SubtitleWsBroadcaster;
  annotationSubtitleWsService: SubtitleWsBroadcaster;
  autoplayReadyGate: Pick<
    ReturnType<typeof createAutoplayReadyGate>,
    | 'maybeSignalPluginAutoplayReady'
    | 'flushPendingAutoplayReadySignal'
    | 'invalidatePendingAutoplayReadyFallbacks'
  >;
  currentMediaTokenizationGate: Pick<
    ReturnType<typeof createCurrentMediaTokenizationGate>,
    'updateCurrentMediaPath' | 'markReady'
  >;
  startupOsdSequencer: Pick<
    ReturnType<typeof createStartupOsdSequencer>,
    | 'reset'
    | 'markTokenizationReady'
    | 'showTokenizationLoading'
    | 'showAnnotationLoading'
    | 'markAnnotationLoadingComplete'
  >;
  managedLocalSubtitleSelectionRuntime: Pick<
    ReturnType<typeof createManagedLocalSubtitleSelectionRuntime>,
    'handleMediaPathChange' | 'handleSubtitleTrackListChange'
  >;
  mediaTimingReviewRuntime: Pick<
    ReturnType<typeof createMediaTimingReviewRuntime>,
    'deferPlaybackResume' | 'cancelPlaybackResume'
  >;
  discordPresenceRuntime: Pick<
    ReturnType<typeof createDiscordPresenceRuntime>,
    'publishDiscordPresence'
  >;
  immersionMediaRuntime: Pick<
    ReturnType<typeof createImmersionMediaRuntime>,
    'syncFromCurrentMediaState'
  >;
  mediaRuntime: Pick<
    ReturnType<typeof createMediaRuntimeService>,
    'updateCurrentMediaPath' | 'updateCurrentMediaTitle'
  >;
  overlayNotificationsRuntime: Pick<
    OverlayNotificationsRuntime,
    | 'showConfiguredStatusNotification'
    | 'showConfiguredPlaybackFeedback'
    | 'getConfiguredStatusNotificationType'
  >;
  overlayModalRuntime: Pick<OverlayModalRuntime, 'handleOverlayModalClosed'>;
  youtube: Pick<YoutubeAssembly, 'isPlaybackActiveNow' | 'shouldQuitWhenMpvCloses'> & {
    playback: Pick<YoutubeAssembly['playback'], 'getQuitOnDisconnectArmed'>;
    primarySubtitleNotification: Pick<
      YoutubeAssembly['primarySubtitleNotification'],
      'handleMediaPathChange' | 'handleSubtitleTrackChange' | 'handleSubtitleTrackListChange'
    >;
    mediaCachePlayback: Pick<YoutubeAssembly['mediaCachePlayback'], 'handleMediaPathChange'>;
    browserPlayback: Pick<
      YoutubeAssembly['browserPlayback'],
      'handleMediaPathChange' | 'handleMpvDisconnected'
    >;
    flow: Pick<
      YoutubeAssembly['flow'],
      | 'handleMediaPathChange'
      | 'cancelWhisperGeneration'
      | 'hasActiveSession'
      | 'cancelActivePicker'
    >;
  };
  jellyfin: Pick<
    JellyfinAssembly,
    | 'getQuitOnDisconnectArmed'
    | 'reportJellyfinRemoteStopped'
    | 'reportJellyfinRemoteProgress'
    | 'getActivePlayback'
    | 'startJellyfinRemoteSession'
  >;
  anilist: Pick<
    AnilistAssembly,
    | 'maybeRunAnilistPostWatchUpdate'
    | 'recordAnilistMediaDuration'
    | 'getCurrentAnilistMediaKey'
    | 'resetAnilistMediaTracking'
    | 'maybeProbeAnilistDuration'
    | 'ensureAnilistMediaGuess'
    | 'resetAnilistMediaGuessState'
  >;
  characterDictionary: {
    autoSync: Pick<CharacterDictionaryAssembly['autoSync'], 'scheduleSync' | 'getCurrentMediaId'>;
    imageLookup: Pick<CharacterDictionaryAssembly['imageLookup'], 'get'>;
    nameCandidates: Pick<CharacterDictionaryAssembly['nameCandidates'], 'get'>;
  };
  dictionaries: {
    jlpt: Pick<DictionariesAssembly['jlpt'], 'ensureJlptDictionaryLookup'>;
    frequency: Pick<DictionariesAssembly['frequency'], 'ensureFrequencyDictionaryLookup'>;
  };
  getSubtitleGenerationRuntime: () => Pick<
    ReturnType<typeof createSubtitleGenerationRuntime>,
    'handleMediaPathChange'
  >;
  /** Plugins observe every mpv client the app creates. */
  getPluginHost: () => { attachMpvClient: (client: MpvIpcClient) => void };
  isCharacterDictionaryEnabled: () => boolean;
  /** Whether the tray icon exists; mpv shutdown keeps the app alive when it does. */
  hasTray: () => boolean;
  loadSubtitleSourceText: typeof loadSubtitleSourceText;
  createSubtitlePrefetchService: (
    serviceDeps: SubtitlePrefetchServiceDeps,
  ) => SubtitlePrefetchService;
  requestAppQuit: () => void;
  maybeStartOverlayLoadingOsd: OverlayNotificationsRuntime['maybeStartOverlayLoadingOsd'];
  flushQueuedMpvOsdNotifications: OverlayNotificationsRuntime['flushQueuedMpvOsdNotifications'];
  refreshMpvSessionBindings: () => Promise<void>;
  ensureImmersionTrackerStarted: () => void;
  restoreOverlayMpvSubtitles: () => void;
  syncOverlayMpvSubtitleSuppression: () => void;
  setOverlayVisible: (visible: boolean) => void;
  ensureYomitanExtensionLoaded: () => Promise<unknown>;
  shouldInitializeMecabForAnnotations: () => boolean;
  getRuntimeBooleanOption: (id: AnnotationRuntimeOptionId, fallback: boolean) => boolean;
  /** Linux fullscreen overlay refresh lives with the overlay window code in `main.ts`. */
  onFullscreenChange: (fullscreen: boolean) => void;
}

export type PlaybackAssembly = ReturnType<typeof createPlaybackAssembly>;

/**
 * Owns the subtitle pipeline around mpv playback: the mpv client runtime and its event
 * handlers, subtitle tokenization and warmups, subtitle prefetch, the subtitle sidebar's cached
 * playback state, secondary subtitles and autoplay subtitle priming. `main.ts` builds it once and
 * starts warmups and the mpv client from its startup code.
 */
export function createPlaybackAssembly(deps: PlaybackAssemblyDeps) {
  const {
    appState,
    configService,
    logger,
    overlayManager,
    subtitleWsService,
    annotationSubtitleWsService,
    autoplayReadyGate,
    currentMediaTokenizationGate,
    startupOsdSequencer,
    managedLocalSubtitleSelectionRuntime,
    mediaTimingReviewRuntime,
    discordPresenceRuntime,
    immersionMediaRuntime,
    mediaRuntime,
    overlayNotificationsRuntime,
    youtube,
    jellyfin,
    anilist,
    characterDictionary,
    dictionaries,
    loadSubtitleSourceText,
    createSubtitlePrefetchService,
  } = deps;
  let backgroundWarmupsStarted = false;
  let tokenizeSubtitleDeferred: ((text: string) => Promise<SubtitleData>) | null = null;
  // Set from the renderer's modal open/close reports; replacement overlay windows restore it.
  let subtitleSidebarRequestedOpen = false;

  function resetSubtitleSidebarEmbeddedLayout(): void {
    sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'video-margin-ratio-right', 0]);
    sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'video-pan-x', 0]);
  }

  function withCurrentSubtitleTiming(payload: SubtitleData): SubtitleData {
    const canonical = resolveCanonicalPrimarySubtitle({
      liveText: payload.text,
      currentTimeSec: toSubtitleClock(
        Number(appState.mpvClient?.currentTimePos),
        appState.mpvClient?.currentSubDelay ?? 0,
      ),
      cues: appState.activeParsedSubtitleCues,
    });
    return {
      ...payload,
      startTime: canonical?.startTime ?? appState.mpvClient?.currentSubStart ?? null,
      endTime: canonical?.endTime ?? appState.mpvClient?.currentSubEnd ?? null,
    };
  }

  function emitSubtitlePayload(
    payload: SubtitleData,
    options?: { resumePrefetch?: boolean },
  ): void {
    const timedPayload = withCurrentSubtitleTiming(payload);
    const currentSubtitleData = appState.currentSubtitleData;
    const isAnnotationUpgrade = isSubtitleAnnotationUpgrade(currentSubtitleData, timedPayload);
    const frequencyDictionary = configService.getConfig().subtitleStyle.frequencyDictionary;
    const frequencyOptions = {
      enabled: frequencyDictionary.enabled,
      topX: frequencyDictionary.topX,
      mode: frequencyDictionary.mode,
    };
    appState.currentSubtitleData = timedPayload;
    overlayManager.broadcastToOverlayWindows('subtitle:set', timedPayload);
    if (!isAnnotationUpgrade) {
      subtitleWsService.broadcast(timedPayload, frequencyOptions);
    }
    annotationSubtitleWsService.broadcast(timedPayload, frequencyOptions);
    autoplayReadyGate.maybeSignalPluginAutoplayReady(timedPayload, { forceWhilePaused: true });
    // resumePrefetch: false marks an emit that is not the end of the work for
    // this line; prefetch stays paused until the subtitle processing controller
    // settles so it does not compete with the on-screen line for the single
    // Yomitan parser window.
    if (options?.resumePrefetch !== false) {
      subtitlePrefetchService?.resume();
    }
  }
  function getCurrentAutoplaySubtitlePayload(): SubtitleData | null {
    const payload = appState.currentSubtitleData;
    if (!payload?.text.trim()) {
      return null;
    }
    if (payload.text !== appState.currentSubText) {
      return null;
    }
    return payload;
  }

  function signalCurrentSubtitleAutoplayReady(): void {
    autoplayReadyGate.flushPendingAutoplayReadySignal();
    const payload = getCurrentAutoplaySubtitlePayload();
    if (payload) {
      autoplayReadyGate.maybeSignalPluginAutoplayReady(payload, { forceWhilePaused: true });
      return;
    }
    if (!appState.currentSubText.trim()) {
      autoplayReadyGate.maybeSignalPluginAutoplayReady(
        { text: '__warm__', tokens: null },
        { forceWhilePaused: true },
      );
    }
  }
  const subtitleProcessingController = createPrefetchAwareSubtitleProcessingController({
    tokenizeSubtitle: async (text: string) =>
      tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : { text, tokens: null },
    emitSubtitlePayload: (payload, options) => emitSubtitlePayload(payload, options),
    resumePrefetch: () => {
      subtitlePrefetchService?.resume();
    },
    logDebug: (message) => {
      logger.debug(`[subtitle-processing] ${message}`);
    },
    now: () => Date.now(),
  });

  let subtitlePrefetchService: SubtitlePrefetchService | null = null;
  let lastObservedTimePos = 0;
  let lastObservedPrimarySubtitleTrackId: number | null = null;
  let subtitleSidebarCuesRef: SubtitleCue[] | null = null;
  let subtitleSidebarCuesRevision = 0;

  /**
   * Cheap sidebar state polled ~12x/s while the sidebar is open. Reads only cached state, no
   * mpv round-trips. `cuesRevision` bumps whenever a new parsed cue array is swapped in, so a
   * re-parse of the same source still tells the renderer to fetch the full snapshot.
   */
  function getSubtitleSidebarPlaybackState(): SubtitleSidebarPlaybackState {
    if (appState.activeParsedSubtitleCues !== subtitleSidebarCuesRef) {
      subtitleSidebarCuesRef = appState.activeParsedSubtitleCues;
      subtitleSidebarCuesRevision += 1;
    }
    const subDelay = appState.mpvClient?.currentSubDelay ?? 0;
    return {
      sourceKey: JSON.stringify([
        appState.activeParsedSubtitleMediaPath,
        appState.activeParsedSubtitleSource,
      ]),
      cuesRevision: subtitleSidebarCuesRevision,
      currentTimeSec:
        appState.mpvClient?.currentTimePos === undefined
          ? null
          : toSubtitleClock(appState.mpvClient.currentTimePos, subDelay),
      subDelay,
      currentSubtitle: {
        text: appState.currentSubText,
        startTime: appState.mpvClient?.currentSubStart ?? null,
        endTime: appState.mpvClient?.currentSubEnd ?? null,
      },
    };
  }
  const SEEK_THRESHOLD_SECONDS = 3;
  const EXPLICIT_SEEK_INTENT_TTL_MS = 2000;
  let explicitSeekIntentExpiresAtMs = 0;

  function isExplicitMpvSeekCommand(command: readonly (string | number)[]): boolean {
    return command[0] === 'seek' || command[0] === 'sub-seek';
  }

  function getMpvPauseCommandValue(command: readonly (string | number)[]): 'yes' | 'no' | null {
    if ((command[0] !== 'set_property' && command[0] !== 'set') || command[1] !== 'pause') {
      return null;
    }
    return command[2] === 'yes' || command[2] === 'no' ? command[2] : null;
  }

  function sendRendererMpvCommand(rawCommand: (string | number)[]): void {
    // Overlay auto-pause releases (popup closed, hover left) must not resume playback
    // behind an open timing review; the review applies them when it closes. A pause
    // request during the review keeps playback paused after it closes.
    const pauseValue = getMpvPauseCommandValue(rawCommand);
    if (pauseValue === 'no' && mediaTimingReviewRuntime.deferPlaybackResume()) {
      return;
    }
    if (pauseValue === 'yes') {
      mediaTimingReviewRuntime.cancelPlaybackResume();
    }
    const command =
      resolveSanitizedSubtitleSeekCommand(
        rawCommand,
        appState.activeParsedSubtitleCues,
        appState.mpvClient?.currentTimePos ?? Number.NaN,
        appState.mpvClient?.currentSubDelay ?? 0,
      ) ?? rawCommand;
    if (isExplicitMpvSeekCommand(command)) {
      explicitSeekIntentExpiresAtMs = Date.now() + EXPLICIT_SEEK_INTENT_TTL_MS;
    }
    sendMpvCommandRuntime(appState.mpvClient, command);
  }

  function consumeExplicitSeekIntent(): boolean {
    const pending = explicitSeekIntentExpiresAtMs >= Date.now();
    explicitSeekIntentExpiresAtMs = 0;
    return pending;
  }

  const autoplaySubtitlePrimingRuntime = createAutoplaySubtitlePrimingRuntime({
    getCurrentMediaPath: () => appState.currentMediaPath,
    getMpvClient: () => appState.mpvClient,
    setCurrentSubText: (text) => {
      appState.currentSubText = text;
    },
    getCurrentSubText: () => appState.currentSubText,
    getCurrentSubtitleData: () => appState.currentSubtitleData,
    getActiveParsedSubtitleCues: () => appState.activeParsedSubtitleCues,
    setActiveParsedSubtitleMediaPath: (mediaPath) => {
      appState.activeParsedSubtitleMediaPath = mediaPath;
    },
    subtitleProcessingController,
    emitSubtitlePayload: (payload, options) => emitSubtitlePayload(payload, options),
    getSubtitlePrefetchService: () => subtitlePrefetchService,
    getLastObservedTimePos: () => lastObservedTimePos,
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    emitSecondarySubtitle: (text) => {
      secondarySubtitleTrackController.handleLiveText(text);
    },
    initSubtitlePrefetch: (sourcePath, currentTimePos, sourceKey) =>
      subtitlePrefetchInitController.initSubtitlePrefetch(sourcePath, currentTimePos, sourceKey),
    refreshSubtitlePrefetchFromActiveTrack: () => refreshSubtitlePrefetchFromActiveTrackHandler(),
    logDebug: (message) => {
      logger.debug(message);
    },
  });

  const subtitlePrefetchInitController = createSubtitlePrefetchInitController({
    getCurrentService: () => subtitlePrefetchService,
    setCurrentService: (service) => {
      subtitlePrefetchService = service;
    },
    loadSubtitleSourceText,
    parseSubtitleCues: (content, filename) => parseSubtitleCues(content, filename),
    createSubtitlePrefetchService: (serviceDeps) => createSubtitlePrefetchService(serviceDeps),
    tokenizeSubtitle: async (text) =>
      tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
    preCacheTokenization: (text, data) => {
      subtitleProcessingController.preCacheTokenization(text, data);
    },
    hasCachedTokenization: (text) => subtitleProcessingController.hasCachedSubtitle(text),
    getCacheGeneration: () => subtitleProcessingController.getCacheGeneration(),
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
    onParsedSubtitleCuesChanged: (cues, sourceKey) => {
      appState.activeParsedSubtitleCues = cues ?? [];
      appState.activeParsedSubtitleSource = sourceKey;
      if (!cues?.length) {
        appState.activeParsedSubtitleMediaPath = null;
      }
      const mediaPath = autoplaySubtitlePrimingRuntime.getCurrentAutoplayMediaPath();
      if (mediaPath && cues?.length) {
        void autoplaySubtitlePrimingRuntime
          .primeAutoplaySubtitleFromParsedCues(mediaPath, cues)
          .catch((error) => {
            logger.debug(
              `[autoplay-subtitle-prime] failed to prime from parsed cues: ${
                error instanceof Error ? error.message : String(error)
              }`,
            );
          });
      }
    },
  });
  const cachedInternalSubtitleTrackExtractor = createCachedInternalSubtitleTrackExtractor();
  const detectRemoteMediaPath = createRemoteMediaPathDetector();
  const resolveActiveSubtitleSidebarSourceHandler = createResolveActiveSubtitleSidebarSourceHandler(
    {
      getFfmpegPath: () => configService.getConfig().subsync.ffmpeg_path.trim() || 'ffmpeg',
      extractInternalSubtitleTrack: (ffmpegPath, videoPath, track) =>
        cachedInternalSubtitleTrackExtractor.extract(ffmpegPath, videoPath, track),
      logDebug: (message) => logger.debug(message),
    },
  );

  const secondarySubtitleTrackController = createSecondarySubtitleTrackController({
    getMpvClient: () => appState.mpvClient,
    getCurrentTimePos: () => appState.mpvClient?.currentTimePos ?? lastObservedTimePos,
    resolveSubtitleSource: (input) => resolveActiveSubtitleSidebarSourceHandler(input),
    loadSubtitleSourceText,
    parseSubtitleCues: (content, filename) => parseSubtitleCues(content, filename),
    setCurrentSecondaryText: (text) => {
      if (appState.mpvClient) {
        appState.mpvClient.currentSecondarySubText = text;
      }
    },
    broadcastSecondaryText: (text) => {
      overlayManager.broadcastToOverlayWindows('secondary-subtitle:set', text);
    },
    logDebug: (message) => logger.debug(message),
    logWarn: (message, error) => logger.warn(message, error),
  });

  const refreshSubtitlePrefetchFromActiveTrackHandler =
    createRefreshSubtitlePrefetchFromActiveTrackHandler({
      getMpvClient: () => appState.mpvClient,
      getLastObservedTimePos: () => lastObservedTimePos,
      // Remote media has no extractable on-disk track to fall back to, so a transient
      // resolve miss (sid briefly 'no', a cycle onto an embedded stream track) would
      // otherwise drop a working cue list for the rest of the episode.
      shouldKeepExistingCuesOnMissingSource: async (videoPath) =>
        isYoutubeMediaPath(videoPath) || (await detectRemoteMediaPath(videoPath)),
      subtitlePrefetchInitController,
      resolveActiveSubtitleSidebarSource: (input) =>
        resolveActiveSubtitleSidebarSourceHandler(input),
      logDebug: (message) => logger.debug(message),
      logWarn: (message) => logger.warn(message),
    });

  const subtitlePrefetchRuntime = {
    cancelPendingInit: () => subtitlePrefetchInitController.cancelPendingInit(),
    initSubtitlePrefetch: subtitlePrefetchInitController.initSubtitlePrefetch,
    refreshSubtitleSidebarFromSource: (sourcePath: string, mediaPath?: string) =>
      autoplaySubtitlePrimingRuntime.refreshSubtitleSidebarFromSource(sourcePath, mediaPath),
    refreshSubtitlePrefetchFromActiveTrack: () => refreshSubtitlePrefetchFromActiveTrackHandler(),
    scheduleSubtitlePrefetchRefresh: (delayMs?: number) =>
      autoplaySubtitlePrimingRuntime.scheduleSubtitlePrefetchRefresh(delayMs),
    clearScheduledSubtitlePrefetchRefresh: () =>
      autoplaySubtitlePrimingRuntime.clearScheduledSubtitlePrefetchRefresh(),
  } as const;

  function refreshCurrentSubtitleAnnotations(): void {
    const hasCurrentSubtitle = appState.currentSubText.trim().length > 0;
    if (hasCurrentSubtitle) {
      subtitlePrefetchService?.pause();
    }
    subtitleProcessingController.invalidateTokenizationCache();
    subtitlePrefetchService?.onSeek(lastObservedTimePos);
    if (!subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText)) {
      // Idle controller: no settle is coming to release the pause above.
      subtitlePrefetchService?.resume();
    }
  }

  let signalAutoplayReadyFromWarmTokenization: ((path: string | null | undefined) => void) | null =
    null;

  const {
    bindMpvClientEventHandlers,
    createMpvClientRuntimeService: createMpvClientRuntimeServiceHandler,
    updateMpvSubtitleRenderMetrics: updateMpvSubtitleRenderMetricsHandler,
    tokenizeSubtitle: tokenizeSubtitleRuntime,
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
      appState,
      getQuitOnDisconnectArmed: () =>
        jellyfin.getQuitOnDisconnectArmed() || youtube.playback.getQuitOnDisconnectArmed(),
      scheduleQuitCheck: (callback) => {
        setTimeout(callback, 500);
      },
      quitApp: () => deps.requestAppQuit(),
      reportJellyfinRemoteStopped: () => {
        void jellyfin.reportJellyfinRemoteStopped();
      },
      onMpvConnected: () => {
        deps.maybeStartOverlayLoadingOsd();
        deps.flushQueuedMpvOsdNotifications();
        secondarySubtitleTrackController.scheduleRefresh(0);
        void deps.refreshMpvSessionBindings();
        if (appState.sessionBindingsInitialized) {
          sendMpvCommandRuntime(appState.mpvClient, [
            'script-message',
            'subminer-reload-session-bindings',
          ]);
        }
        if (appState.currentSubText.trim()) {
          subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
        }
      },
      maybeRunAnilistPostWatchUpdate: (options) => anilist.maybeRunAnilistPostWatchUpdate(options),
      recordAnilistMediaDuration: (durationSec) => {
        anilist.recordAnilistMediaDuration(durationSec);
      },
      logSubtitleTimingError: (message, error) => logger.error(message, error),
      broadcastToOverlayWindows: (channel, payload) => {
        overlayManager.broadcastToOverlayWindows(channel, payload);
      },
      onSecondarySubtitleChange: (text) => {
        secondarySubtitleTrackController.handleLiveText(text);
      },
      getImmediateSubtitlePayload: (text) =>
        subtitleProcessingController.consumeCachedSubtitle(text),
      emitImmediateSubtitle: (payload) => {
        emitSubtitlePayload(payload);
      },
      onSubtitleChange: (text) => {
        // Pause only; restarting the prefetch run here would discard in-flight
        // tokenization work on every line. Real seeks restart via onTimePosUpdate.
        subtitlePrefetchService?.pause();
        if (!subtitleProcessingController.onSubtitleChange(text)) {
          // Repeat of the current text: the controller is idle, so no settle is
          // coming to release the pause. Resume now instead of idling prefetch
          // for the rest of the cue.
          subtitlePrefetchService?.resume();
        }
      },
      refreshDiscordPresence: () => {
        discordPresenceRuntime.publishDiscordPresence();
      },
      logSubtitleProcessingDebug: (message: string) => logger.debug(message),
      ensureImmersionTrackerInitialized: () => {
        deps.ensureImmersionTrackerStarted();
      },
      tokenizeSubtitleForImmersion: async (text): Promise<SubtitleData | null> =>
        tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
      updateCurrentMediaPath: (path) => {
        const normalizedPath = path.trim();
        deps.maybeStartOverlayLoadingOsd(normalizedPath);
        const previousPath = appState.currentMediaPath?.trim() || null;
        const preserveParsedSubtitleCues = isSameYoutubeMediaPath(
          normalizedPath,
          appState.activeParsedSubtitleMediaPath,
        );
        if ((normalizedPath || null) !== previousPath) {
          cachedInternalSubtitleTrackExtractor.clear();
          secondarySubtitleTrackController.reset();
          const resetSubtitlePayload = { text: '', tokens: null };
          const frequencyDictionary = configService.getConfig().subtitleStyle.frequencyDictionary;
          const frequencyOptions = {
            enabled: frequencyDictionary.enabled,
            topX: frequencyDictionary.topX,
            mode: frequencyDictionary.mode,
          };
          autoplaySubtitlePrimingRuntime.resetAutoplaySubtitlePrime();
          lastObservedTimePos = 0;
          appState.currentSubText = '';
          appState.currentSubtitleData = null;
          if (!preserveParsedSubtitleCues) {
            appState.activeParsedSubtitleCues = [];
            appState.activeParsedSubtitleSource = null;
            appState.activeParsedSubtitleMediaPath = null;
          }
          overlayManager.broadcastToOverlayWindows('subtitle:set', resetSubtitlePayload);
          subtitleWsService.broadcast(resetSubtitlePayload, frequencyOptions);
          annotationSubtitleWsService.broadcast(resetSubtitlePayload, frequencyOptions);
          autoplayReadyGate.invalidatePendingAutoplayReadyFallbacks();
        }
        currentMediaTokenizationGate.updateCurrentMediaPath(path);
        managedLocalSubtitleSelectionRuntime.handleMediaPathChange(path);
        startupOsdSequencer.reset();
        subtitlePrefetchRuntime.clearScheduledSubtitlePrefetchRefresh();
        if (!preserveParsedSubtitleCues) {
          subtitlePrefetchRuntime.cancelPendingInit();
        }
        youtube.primarySubtitleNotification.handleMediaPathChange(path);
        void youtube.mediaCachePlayback.handleMediaPathChange(path);
        youtube.browserPlayback.handleMediaPathChange(path);
        youtube.flow.handleMediaPathChange(path);
        deps.getSubtitleGenerationRuntime().handleMediaPathChange(path);
        if (path) {
          deps.ensureImmersionTrackerStarted();
          secondarySubtitleTrackController.scheduleRefresh();
          void subtitlePrefetchRuntime.refreshSubtitlePrefetchFromActiveTrack();
          // Retry after a short delay because MPV can populate track-list after path.
          subtitlePrefetchRuntime.scheduleSubtitlePrefetchRefresh(500);
        }
        mediaRuntime.updateCurrentMediaPath(path);
      },
      restoreMpvSubVisibility: () => {
        deps.restoreOverlayMpvSubtitles();
      },
      resetSubtitleSidebarEmbeddedLayout: () => {
        resetSubtitleSidebarEmbeddedLayout();
      },
      getCurrentAnilistMediaKey: () => anilist.getCurrentAnilistMediaKey(),
      resetAnilistMediaTracking: (mediaKey) => {
        anilist.resetAnilistMediaTracking(mediaKey);
      },
      maybeProbeAnilistDuration: (mediaKey) => {
        void anilist.maybeProbeAnilistDuration(mediaKey);
      },
      ensureAnilistMediaGuess: (mediaKey) => {
        void anilist.ensureAnilistMediaGuess(mediaKey);
      },
      syncImmersionMediaState: () => {
        immersionMediaRuntime.syncFromCurrentMediaState();
      },
      signalAutoplayReadyIfWarm: (path) => signalAutoplayReadyFromWarmTokenization?.(path),
      markJellyfinRemotePlaybackLoaded: (path) => {
        markJellyfinRemotePlaybackLoadedState(jellyfin.getActivePlayback(), path);
      },
      scheduleCharacterDictionarySync: () => {
        if (
          !configService.getConfig().subtitleStyle.nameMatchEnabled ||
          !deps.isCharacterDictionaryEnabled() ||
          youtube.isPlaybackActiveNow()
        ) {
          return;
        }
        characterDictionary.autoSync.scheduleSync();
      },
      updateCurrentMediaTitle: (title) => {
        mediaRuntime.updateCurrentMediaTitle(title);
      },
      resetAnilistMediaGuessState: () => {
        anilist.resetAnilistMediaGuessState();
      },
      reportJellyfinRemoteProgress: (forceImmediate) => {
        void jellyfin.reportJellyfinRemoteProgress(forceImmediate);
      },
      consumeExplicitSeek: () => consumeExplicitSeekIntent(),
      onTimePosUpdate: (time) => {
        const delta = time - lastObservedTimePos;
        if (subtitlePrefetchService && (delta > SEEK_THRESHOLD_SECONDS || delta < 0)) {
          subtitlePrefetchService.onSeek(time);
        }
        lastObservedTimePos = time;
        secondarySubtitleTrackController.handleTimePos(time);
      },
      onFullscreenChange: (fullscreen) => deps.onFullscreenChange(fullscreen),
      onSubtitleTrackChange: (sid) => {
        lastObservedPrimarySubtitleTrackId = sid;
        logger.info('[mpv-subtitles] primary subtitle track changed', { sid });
        autoplaySubtitlePrimingRuntime.scheduleSubtitlePrefetchRefresh();
        youtube.primarySubtitleNotification.handleSubtitleTrackChange(sid);
      },
      onSecondarySubtitleTrackChange: () => {
        secondarySubtitleTrackController.handleTrackChange();
        secondarySubtitleTrackController.scheduleRefresh(0);
      },
      onSecondarySubtitleDelayChange: (delay) => {
        secondarySubtitleTrackController.handleDelayChange(delay);
      },
      onSubtitleTrackListChange: (trackList) => {
        const diagnostics = buildSubtitleTrackDiagnostics(
          lastObservedPrimarySubtitleTrackId,
          trackList,
        );
        if (!diagnostics.trackListReadable) {
          logger.warn(
            '[mpv-subtitles] mpv reported an unreadable subtitle track list',
            diagnostics,
          );
        } else if (diagnostics.subtitleTrackCount === 0) {
          logger.warn('[mpv-subtitles] mpv reported no subtitle tracks', diagnostics);
        } else {
          logger.info('[mpv-subtitles] subtitle track list updated', diagnostics);
        }
        managedLocalSubtitleSelectionRuntime.handleSubtitleTrackListChange(trackList);
        secondarySubtitleTrackController.scheduleRefresh(0);
        autoplaySubtitlePrimingRuntime.scheduleSubtitlePrefetchRefresh();
        youtube.primarySubtitleNotification.handleSubtitleTrackListChange(trackList);
      },
      updateSubtitleRenderMetrics: (patch) => {
        updateMpvSubtitleRenderMetricsHandler(patch as Partial<MpvSubtitleRenderMetrics>);
      },
      syncOverlayMpvSubtitleSuppression: () => {
        deps.syncOverlayMpvSubtitleSuppression();
      },
    },
    mpvClientRuntimeServiceFactoryMainDeps: {
      createClient: MpvIpcClient,
      getSocketPath: () => appState.mpvSocketPath,
      getResolvedConfig: () => configService.getConfig(),
      isAutoStartOverlayEnabled: () => appState.autoStartOverlay,
      setOverlayVisible: (visible: boolean) => deps.setOverlayVisible(visible),
      isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      getReconnectTimer: () => appState.reconnectTimer,
      setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => {
        appState.reconnectTimer = timer;
      },
      shouldAutoLoadSecondarySubTrack: (path: string) =>
        shouldAutoLoadSecondarySubTrackForJellyfinPlayback(jellyfin.getActivePlayback(), path),
      shouldQuitOnMpvShutdown: () =>
        shouldQuitOnMpvShutdownForTrayState({
          managedPlayback: appState.initialArgs?.managedPlayback === true,
          backgroundMode: appState.backgroundMode,
          hasTray: deps.hasTray(),
        }),
      requestAppQuit: () => deps.requestAppQuit(),
    },
    updateMpvSubtitleRenderMetricsMainDeps: {
      getCurrentMetrics: () => appState.mpvSubtitleRenderMetrics,
      setCurrentMetrics: (metrics) => {
        appState.mpvSubtitleRenderMetrics = metrics;
      },
      applyPatch: (current, patch) => applyMpvSubtitleRenderMetricsPatch(current, patch),
      broadcastMetrics: () => {
        // no renderer consumer for subtitle render metrics updates at present
      },
    },
    tokenizer: {
      buildTokenizerDepsMainDeps: {
        getYomitanExt: () => appState.yomitanExt,
        getYomitanSession: () => appState.yomitanSession,
        getYomitanParserWindow: () => appState.yomitanParserWindow,
        setYomitanParserWindow: (window) => {
          appState.yomitanParserWindow = window as BrowserWindow | null;
        },
        getYomitanParserReadyPromise: () => appState.yomitanParserReadyPromise,
        setYomitanParserReadyPromise: (promise) => {
          appState.yomitanParserReadyPromise = promise;
        },
        getYomitanParserInitPromise: () => appState.yomitanParserInitPromise,
        setYomitanParserInitPromise: (promise) => {
          appState.yomitanParserInitPromise = promise;
        },
        isKnownWord: (text, reading, options) =>
          Boolean(appState.ankiIntegration?.isKnownWord(text, reading, options)),
        getKnownWordTier: (text, reading, options) =>
          appState.ankiIntegration?.getKnownWordTier(text, reading, options) ?? null,
        recordLookup: (hit) => {
          deps.ensureImmersionTrackerStarted();
          appState.immersionTracker?.recordLookup(hit);
        },
        getKnownWordMatchMode: () =>
          appState.ankiIntegration?.getKnownWordMatchMode() ??
          configService.getConfig().ankiConnect.knownWords.matchMode,
        getKnownWordsEnabled: () =>
          deps.getRuntimeBooleanOption(
            'subtitle.annotation.knownWords.highlightEnabled',
            configService.getConfig().ankiConnect.knownWords.highlightEnabled,
          ),
        getNPlusOneEnabled: () =>
          deps.getRuntimeBooleanOption(
            'subtitle.annotation.nPlusOne',
            configService.getConfig().ankiConnect.nPlusOne.enabled,
          ),
        getMinSentenceWordsForNPlusOne: () =>
          configService.getConfig().ankiConnect.nPlusOne.minSentenceWords,
        getJlptLevel: (text) => appState.jlptLevelLookup(text),
        getJlptEnabled: () =>
          deps.getRuntimeBooleanOption(
            'subtitle.annotation.jlpt',
            configService.getConfig().subtitleStyle.enableJlpt,
          ),
        getCharacterDictionaryEnabled: () =>
          configService.getConfig().subtitleStyle.nameMatchEnabled &&
          deps.isCharacterDictionaryEnabled() &&
          !youtube.isPlaybackActiveNow(),
        getNameMatchEnabled: () => configService.getConfig().subtitleStyle.nameMatchEnabled,
        getNameMatchImagesEnabled: () =>
          configService.getConfig().subtitleStyle.nameMatchImagesEnabled,
        getCharacterNameImage: (term) => characterDictionary.imageLookup.get(term),
        getCurrentCharacterDictionaryMediaId: () =>
          characterDictionary.autoSync.getCurrentMediaId(),
        getCharacterNameCandidates: () => characterDictionary.nameCandidates.get(),
        getFrequencyDictionaryEnabled: () =>
          deps.getRuntimeBooleanOption(
            'subtitle.annotation.frequency',
            configService.getConfig().subtitleStyle.frequencyDictionary.enabled,
          ),
        getFrequencyDictionaryMatchMode: () =>
          configService.getConfig().subtitleStyle.frequencyDictionary.matchMode,
        getFrequencyRank: (text) => appState.frequencyRankLookup(text),
        getYomitanGroupDebugEnabled: () => appState.overlayDebugVisualizationEnabled,
        getMecabTokenizer: () => appState.mecabTokenizer,
        onTokenizationReady: () => {
          currentMediaTokenizationGate.markReady(
            appState.currentMediaPath?.trim() ||
              appState.mpvClient?.currentVideoPath?.trim() ||
              null,
          );
          startupOsdSequencer.markTokenizationReady();
        },
      },
      createTokenizerRuntimeDeps: (tokenizerDeps) => createTokenizerDepsRuntime(tokenizerDeps),
      tokenizeSubtitle: (text, tokenizerDeps) => tokenizeSubtitleCore(text, tokenizerDeps),
      createMecabTokenizerAndCheckMainDeps: {
        getMecabTokenizer: () => appState.mecabTokenizer,
        setMecabTokenizer: (tokenizer) => {
          appState.mecabTokenizer = tokenizer as MecabTokenizer | null;
        },
        createMecabTokenizer: () => new MecabTokenizer(),
        checkAvailability: async (tokenizer) => (tokenizer as MecabTokenizer).checkAvailability(),
      },
      prewarmSubtitleDictionariesMainDeps: {
        ensureJlptDictionaryLookup: () => dictionaries.jlpt.ensureJlptDictionaryLookup(),
        ensureFrequencyDictionaryLookup: () =>
          dictionaries.frequency.ensureFrequencyDictionaryLookup(),
        showMpvOsd: (message: string) =>
          overlayNotificationsRuntime.showConfiguredStatusNotification(message),
        showLoadingOsd: (message: string) => startupOsdSequencer.showAnnotationLoading(message),
        showLoadedOsd: (message: string) =>
          startupOsdSequencer.markAnnotationLoadingComplete(message),
        shouldShowOsdNotification: () => {
          const type = overlayNotificationsRuntime.getConfiguredStatusNotificationType();
          return type === 'osd' || type === 'osd-system';
        },
      },
    },
    warmups: {
      launchBackgroundWarmupTaskMainDeps: {
        now: () => Date.now(),
        logDebug: (message) => logger.debug(message),
        logWarn: (message) => logger.warn(message),
      },
      startBackgroundWarmupsMainDeps: {
        getStarted: () => backgroundWarmupsStarted,
        setStarted: (started) => {
          backgroundWarmupsStarted = started;
        },
        isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
        ensureYomitanExtensionLoaded: () => deps.ensureYomitanExtensionLoaded().then(() => {}),
        shouldWarmupMecab: () => {
          const startupWarmups = configService.getConfig().startupWarmups;
          if (startupWarmups.lowPowerMode) {
            return false;
          }
          if (!startupWarmups.mecab) {
            return false;
          }
          return deps.shouldInitializeMecabForAnnotations();
        },
        shouldWarmupYomitanExtension: () =>
          configService.getConfig().startupWarmups.yomitanExtension,
        shouldWarmupSubtitleDictionaries: () => {
          const startupWarmups = configService.getConfig().startupWarmups;
          if (startupWarmups.lowPowerMode) {
            return false;
          }
          return startupWarmups.subtitleDictionaries;
        },
        shouldWarmupJellyfinRemoteSession: () => {
          const startupWarmups = configService.getConfig().startupWarmups;
          if (startupWarmups.lowPowerMode) {
            return false;
          }
          return startupWarmups.jellyfinRemoteSession;
        },
        shouldAutoConnectJellyfinRemote: () => {
          const jellyfin = configService.getConfig().jellyfin;
          return (
            jellyfin.enabled && jellyfin.remoteControlEnabled && jellyfin.remoteControlAutoConnect
          );
        },
        startJellyfinRemoteSession: () => jellyfin.startJellyfinRemoteSession(),
        logDebug: (message) => logger.debug(message),
      },
    },
  });

  async function tokenizeSubtitle(text: string): Promise<SubtitleData> {
    if (!isTokenizationWarmupReady()) {
      startupOsdSequencer.showTokenizationLoading('Loading subtitle tokenization...');
    }
    return await tokenizeSubtitleRuntime(text);
  }

  signalAutoplayReadyFromWarmTokenization = createAutoplayTokenizationWarmRelease({
    isTokenizationWarmupReady: () => isTokenizationWarmupReady(),
    startTokenizationWarmups: async () => {
      await startTokenizationWarmups();
    },
    getCurrentMediaPath: () =>
      appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
    primeCurrentSubtitle: (mediaPath) =>
      autoplaySubtitlePrimingRuntime.primeCurrentSubtitleForAutoplay(mediaPath),
    signalAutoplayReady: () => signalCurrentSubtitleAutoplayReady(),
    warn: (message, error) => logger.warn(message, error),
  });
  tokenizeSubtitleDeferred = tokenizeSubtitle;

  const aniSkipRuntime = createAniSkipRuntime({
    getAniSkipConfig: () => ({
      aniskipEnabled: configService.getConfig().mpv.aniskipEnabled,
      aniskipButtonKey: configService.getConfig().mpv.aniskipButtonKey,
    }),
    resolveMetadataForFile: (mediaPath) => resolveAniSkipMetadataForFile(mediaPath),
    sendMpvCommand: (command) => {
      appState.mpvClient?.send({ command });
    },
    requestMpvProperty: (name) => {
      if (!appState.mpvClient) {
        return Promise.reject(new Error('MPV not connected'));
      }
      return appState.mpvClient.requestProperty(name);
    },
    isMpvConnected: () => appState.mpvClient?.connected === true,
    getCurrentTimePos: () => appState.mpvClient?.currentTimePos ?? Number.NaN,
    showMpvOsd: (text, durationMs) => {
      appState.mpvClient?.send({ command: ['show-text', text, durationMs] });
    },
    showPlaybackFeedback: (text) =>
      overlayNotificationsRuntime.showConfiguredPlaybackFeedback(text),
    logInfo: (message) => logger.info(message),
    logWarn: (message, error) => logger.warn(message, error),
    logDebug: (message) => logger.debug(message),
  });

  /** Creates the app's mpv client with YouTube, AniSkip and plugin listeners attached. */
  function createMpvClient(): MpvIpcClient {
    const client = createMpvClientRuntimeServiceHandler() as MpvIpcClient;
    client.on('connection-change', ({ connected }) => {
      if (connected) {
        return;
      }
      youtube.browserPlayback.handleMpvDisconnected();
      youtube.flow.cancelWhisperGeneration();
      deps.getSubtitleGenerationRuntime().handleMediaPathChange(null);
      if (youtube.shouldQuitWhenMpvCloses()) {
        deps.requestAppQuit();
        return;
      }
      if (!youtube.flow.hasActiveSession()) {
        return;
      }
      youtube.flow.cancelActivePicker();
      overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.youtubePickerCancel, null);
      deps.overlayModalRuntime.handleOverlayModalClosed('youtube-track-picker');
    });
    client.on('connection-change', aniSkipRuntime.handleConnectionChange);
    client.on('media-path-change', aniSkipRuntime.handleMediaPathChange);
    client.on('time-pos-change', aniSkipRuntime.handleTimePosChange);
    client.on('client-message', aniSkipRuntime.handleClientMessage);
    deps.getPluginHost().attachMpvClient(client);
    return client;
  }

  const ipc: IpcSlice<
    | 'tokenizeCurrentSubtitle'
    | 'getCurrentSubtitleRaw'
    | 'getSubtitleSidebarOpen'
    | 'getSubtitleSidebarPlaybackState'
    | 'getSubtitleSidebarSnapshot'
    | 'getPlaybackPaused'
    | 'getMpvClient'
  > = {
    tokenizeCurrentSubtitle: async () => {
      const tokenizeSubtitleForCurrent = tokenizeSubtitleDeferred;
      return resolveCurrentSubtitleForRenderer({
        currentSubText: appState.currentSubText,
        currentSubtitleData: appState.currentSubtitleData,
        withCurrentSubtitleTiming: (payload) => withCurrentSubtitleTiming(payload),
        tokenizeUncached: false,
        tokenizeSubtitle: tokenizeSubtitleForCurrent
          ? (text) => tokenizeSubtitleForCurrent(text)
          : undefined,
        onResolvedSubtitle: (payload) => {
          appState.currentSubtitleData = payload;
          autoplayReadyGate.maybeSignalPluginAutoplayReady(payload, { forceWhilePaused: true });
        },
      });
    },
    getCurrentSubtitleRaw: () => appState.currentSubText,
    getSubtitleSidebarOpen: () => subtitleSidebarRequestedOpen,
    getSubtitleSidebarPlaybackState: () => getSubtitleSidebarPlaybackState(),
    // Full snapshot: resolves the active source over mpv IPC and re-parses it when it changed.
    // The renderer fetches it on open, on source/cue changes, and on a slow fallback cadence.
    getSubtitleSidebarSnapshot: async () => {
      const config = configService.getConfig().subtitleSidebar;
      const cachedSnapshot = () => ({
        ...getSubtitleSidebarPlaybackState(),
        cues: appState.activeParsedSubtitleCues,
        config,
      });
      const client = appState.mpvClient;
      if (!client?.connected) {
        return cachedSnapshot();
      }

      try {
        const [currentExternalFilenameRaw, currentTrackRaw, trackListRaw, sidRaw, videoPathRaw] =
          await Promise.all([
            client.requestProperty('current-tracks/sub/external-filename').catch(() => null),
            client.requestProperty('current-tracks/sub').catch(() => null),
            client.requestProperty('track-list'),
            client.requestProperty('sid'),
            client.requestProperty('path'),
          ]);
        const videoPath = typeof videoPathRaw === 'string' ? videoPathRaw : '';
        if (!videoPath) {
          return cachedSnapshot();
        }
        if (
          shouldUseCachedYoutubeParsedCues({
            videoPath,
            cachedMediaPath: appState.activeParsedSubtitleMediaPath,
            cachedCueCount: appState.activeParsedSubtitleCues.length,
          })
        ) {
          return cachedSnapshot();
        }

        const resolvedSource = await resolveActiveSubtitleSidebarSourceHandler({
          currentExternalFilenameRaw,
          currentTrackRaw,
          trackListRaw,
          sidRaw,
          videoPath,
        });
        if (!resolvedSource) {
          return cachedSnapshot();
        }

        try {
          if (appState.activeParsedSubtitleSource === resolvedSource.sourceKey) {
            return cachedSnapshot();
          }

          const content = await loadSubtitleSourceText(resolvedSource.path);
          const cues = parseSubtitleCues(content, resolvedSource.path);
          appState.activeParsedSubtitleCues = cues;
          appState.activeParsedSubtitleSource = resolvedSource.sourceKey;
          appState.activeParsedSubtitleMediaPath = videoPath || null;
          return cachedSnapshot();
        } finally {
          await resolvedSource.cleanup?.();
        }
      } catch {
        return cachedSnapshot();
      }
    },
    getPlaybackPaused: () =>
      resolveFreshPlaybackPaused({
        getCachedPlaybackPaused: () => appState.playbackPaused,
        getMpvClient: () => appState.mpvClient,
      }),
    getMpvClient: () => appState.mpvClient,
  };

  return {
    ipc,
    createMpvClient,
    aniSkipRuntime,
    setSubtitleSidebarRequestedOpen: (open: boolean): void => {
      subtitleSidebarRequestedOpen = open;
    },
    subtitleProcessingController,
    autoplaySubtitlePrimingRuntime,
    secondarySubtitleTrackController,
    subtitlePrefetchRuntime,
    cachedInternalSubtitleTrackExtractor,
    resolveActiveSubtitleSidebarSourceHandler,
    withCurrentSubtitleTiming,
    getSubtitleSidebarPlaybackState,
    sendRendererMpvCommand,
    refreshCurrentSubtitleAnnotations,
    tokenizeSubtitle,
    bindMpvClientEventHandlers,
    createMpvClientRuntimeServiceHandler,
    createMecabTokenizerAndCheck,
    prewarmSubtitleDictionaries,
    startBackgroundWarmups,
    startTokenizationWarmups,
    isTokenizationWarmupReady,
    getSubtitlePrefetchService: () => subtitlePrefetchService,
    getLastObservedTimePos: () => lastObservedTimePos,
    getTokenizeSubtitleDeferred: () => tokenizeSubtitleDeferred,
  };
}
