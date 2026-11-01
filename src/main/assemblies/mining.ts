import { resolveMediaGenerationInput } from '../../anki-integration/media-source';
import {
  captureLiveSubtitleMiningContext,
  copyCurrentSubtitle as copyCurrentSubtitleCore,
  createFieldGroupingOverlayRuntime,
  handleMineSentenceDigit as handleMineSentenceDigitCore,
  handleMultiCopyDigit as handleMultiCopyDigitCore,
  markLastCardAsAudioCard as markLastCardAsAudioCardCore,
  mineSentenceCard as mineSentenceCardCore,
  shiftSubtitleRangeToVideoClock,
  toSubtitleClock,
  triggerFieldGrouping as triggerFieldGroupingCore,
  updateLastCardFromClipboard as updateLastCardFromClipboardCore,
  type createOverlayManager,
} from '../../core/services';
import { createMediaTimingFrameExtractor } from '../../core/services/media-timing-frame';
import { MediaTimingPreviewSession } from '../../core/services/media-timing-preview';
import { generateSpeechWaveform } from '../../core/services/media-timing-waveform';
import { getSharedRemoteMediaWindowCache } from '../../core/services/remote-media-window-cache';
import type { Logger } from '../../logger';
import type { PluginMinedNote, PluginMiningProvenance } from '../../plugin-api';
import { IPC_CHANNELS, type OverlayHostedModal } from '../../shared/ipc/contracts';
import type { ResolvedConfig, SubtitleMiningContext } from '../../types';
import type { OverlayModalRuntime } from '../overlay-runtime';
import {
  createCopyCurrentSubtitleHandler,
  createHandleMineSentenceDigitHandler,
  createHandleMultiCopyDigitHandler,
  createMarkLastCardAsAudioCardHandler,
  createMineSentenceCardHandler,
  createRefreshKnownWordCacheHandler,
  createTriggerFieldGroupingHandler,
  createUpdateLastCardFromClipboardHandler,
} from '../runtime/domains/mining';
import {
  createGetFieldGroupingResolverHandler,
  createSetFieldGroupingResolverHandler,
} from '../runtime/domains/overlay';
import {
  collectMediaTimingContextLines,
  createMediaTimingReviewRuntime,
} from '../runtime/media-timing-review';
import { openMediaTimingReviewModal } from '../runtime/media-timing-review-open';
import { findMpvExecutablePath } from '../runtime/mpv-process';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { resolvePrimarySubtitle } from '../runtime/primary-subtitle-text';
import type { AppState } from '../state';
import type { OverlayModalsAssembly } from './overlay-modals';
import type { YoutubeAssembly } from './youtube';

type OverlayManager = ReturnType<typeof createOverlayManager>;
type AnkiIntegrationLike = NonNullable<AppState['ankiIntegration']>;

/**
 * What card mining needs from the rest of the app. Siblings built after this assembly in
 * `main.ts` are reached through callbacks.
 */
export interface MiningAssemblyDeps {
  appState: Pick<
    AppState,
    | 'mpvClient'
    | 'currentMediaPath'
    | 'activeParsedSubtitleCues'
    | 'ankiIntegration'
    | 'subtitleTimingTracker'
    | 'fieldGroupingResolver'
    | 'fieldGroupingResolverSequence'
  >;
  getConfig: () => ResolvedConfig;
  logger: Pick<Logger, 'warn' | 'error'>;
  overlayManager: Pick<
    OverlayManager,
    'getMainWindow' | 'getModalWindow' | 'getVisibleOverlayVisible'
  >;
  overlayModalRuntime: Pick<
    OverlayModalRuntime,
    | 'getRestoreVisibleOverlayOnModalClose'
    | 'waitForModalOpen'
    | 'handleOverlayModalClosed'
    | 'sendToActiveOverlayWindow'
  >;
  youtube: Pick<
    YoutubeAssembly,
    'getCachedMediaPathForCurrentPlayback' | 'shouldRequireMediaCacheForCurrentPlayback'
  >;
  /** Electron's clipboard, injected so the assembly loads under `bun test`. */
  clipboard: { readText: () => string; writeText: (text: string) => void };
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'showConfiguredStatusNotification' | 'showConfiguredPlaybackFeedback'
  >;
  getOverlayModals: () => Pick<OverlayModalsAssembly, 'createOverlayHostedModalOpenDeps'>;
  setVisibleOverlayVisible: (visible: boolean) => void;
  ensureOverlayStartupPrereqs: () => void;
  ensureOverlayWindowsReadyForVisibilityActions: () => void;
  /** Callbacks the Anki integration reports back through once mining finishes. */
  ankiIntegrationCallbacks: {
    recordCardsMined: (count: number, noteIds?: number[]) => void;
    getMiningProvenance: () => PluginMiningProvenance;
    onCardsReady: (noteIds: number[], notes: PluginMinedNote[]) => void;
    onKnownWordCacheUpdated: () => void;
    generateSentenceFurigana: (text: string, highlightedText?: string) => Promise<string | null>;
  };
}

export type MiningAssembly = ReturnType<typeof createMiningAssembly>;

/**
 * Owns card mining: the subtitle context captured for the next card, the mine/copy/update
 * handlers behind shortcuts, CLI and session actions, field grouping, and the media timing
 * review the Anki integration opens before writing media.
 */
export function createMiningAssembly(deps: MiningAssemblyDeps) {
  const { appState, logger, overlayManager, overlayModalRuntime } = deps;
  const showStatus = (text: string) =>
    deps.getOverlayNotifications().showConfiguredStatusNotification(text);
  const showPlaybackFeedback = (text: string) =>
    deps.getOverlayNotifications().showConfiguredPlaybackFeedback(text);
  let pendingSubtitleMiningContext: SubtitleMiningContext | null = null;

  function captureCurrentPrimarySubtitleMiningContext(): SubtitleMiningContext | null {
    // Mine what the overlay shows, not raw mpv `sub-text`: the raw text lists every active
    // event, so a finished caption row lingering beside a fresh line would end up on the
    // card. The parsed view also carries the cue's own timings for the clip range.
    // Cues sit on the subtitle clock, so the lookup undoes `sub-delay` first.
    const subDelay = appState.mpvClient?.currentSubDelay ?? 0;
    const resolved = resolvePrimarySubtitle({
      liveText: appState.mpvClient?.currentSubText ?? '',
      currentTimeSec: toSubtitleClock(Number(appState.mpvClient?.currentTimePos), subDelay),
      cues: appState.activeParsedSubtitleCues,
    });
    // Same validity bar as the live capture path: an unusable resolved span must fall
    // back rather than hand mining an empty line or an inverted range.
    const resolvedText = resolved?.text.replace(/\n{2,}/g, '\n').trim();
    if (
      !resolved ||
      !resolvedText ||
      !Number.isFinite(resolved.startTime) ||
      !Number.isFinite(resolved.endTime) ||
      resolved.endTime <= resolved.startTime
    ) {
      return captureLiveSubtitleMiningContext(appState.mpvClient);
    }
    return shiftSubtitleRangeToVideoClock<SubtitleMiningContext>(
      {
        source: 'overlay',
        text: resolvedText,
        startTime: resolved.startTime,
        endTime: resolved.endTime,
        capturedAtMs: Date.now(),
      },
      subDelay,
    );
  }

  const recordSubtitleMiningContext = (context: SubtitleMiningContext | null): void => {
    pendingSubtitleMiningContext = context;
  };
  /** Hands the recorded context to exactly one card, then forgets it. */
  const consumePendingSubtitleMiningContext = (): SubtitleMiningContext | null => {
    const context = pendingSubtitleMiningContext;
    pendingSubtitleMiningContext = null;
    return context;
  };

  const getFieldGroupingResolver = createGetFieldGroupingResolverHandler({
    getResolver: () => appState.fieldGroupingResolver,
  });

  const setFieldGroupingResolver = createSetFieldGroupingResolverHandler({
    setResolver: (resolver) => {
      appState.fieldGroupingResolver = resolver;
    },
    nextSequence: () => {
      appState.fieldGroupingResolverSequence += 1;
      return appState.fieldGroupingResolverSequence;
    },
    getSequence: () => appState.fieldGroupingResolverSequence,
  });

  const fieldGroupingOverlayRuntime = createFieldGroupingOverlayRuntime<OverlayHostedModal>({
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible) => deps.setVisibleOverlayVisible(visible),
    getResolver: () => getFieldGroupingResolver(),
    setResolver: (resolver) => setFieldGroupingResolver(resolver),
    getRestoreVisibleOverlayOnModalClose: () =>
      overlayModalRuntime.getRestoreVisibleOverlayOnModalClose(),
    // These are optional on the runtime options, so a missing forward compiles silently, but
    // dropping them leaves the field grouping modal with no modal-open ack/retry, no teardown on
    // failure, and no warn logging.
    waitForModalOpen: (modal, timeoutMs) => overlayModalRuntime.waitForModalOpen(modal, timeoutMs),
    handleOverlayModalClosed: (modal) => overlayModalRuntime.handleOverlayModalClosed(modal),
    logWarn: (message) => logger.warn(message),
    ensureOverlayStartupPrereqs: () => deps.ensureOverlayStartupPrereqs(),
    ensureOverlayWindowsReadyForVisibilityActions: () =>
      deps.ensureOverlayWindowsReadyForVisibilityActions(),
    sendToVisibleOverlay: (channel, payload, runtimeOptions) =>
      overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  });

  const sendToOverlayWindows = (channel: string, payload: unknown) => {
    // The review may live in either overlay window; the renderer ignores foreign review ids.
    for (const window of [overlayManager.getMainWindow(), overlayManager.getModalWindow()]) {
      if (window && !window.isDestroyed()) {
        window.webContents.send(channel, payload);
      }
    }
  };
  const getMediaGenerationOptions = () => ({
    getCachedMediaPath: (currentVideoPath: string, kind: 'audio' | 'video') =>
      deps.youtube.getCachedMediaPathForCurrentPlayback(currentVideoPath, kind),
    remoteCacheMode: deps.youtube.shouldRequireMediaCacheForCurrentPlayback()
      ? ('required' as const)
      : ('optional' as const),
  });

  const mediaTimingFrameExtractor = createMediaTimingFrameExtractor();
  const mediaTimingReviewRuntime = createMediaTimingReviewRuntime({
    getMpvClient: () => appState.mpvClient,
    getCurrentMediaPath: () =>
      appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
    // An invalid configured path is passed through so the preview fails instead of using PATH mpv.
    getMpvExecutablePath: () => {
      const configuredPath = deps.getConfig().mpv.executablePath;
      return findMpvExecutablePath(configuredPath) || configuredPath;
    },
    createPreviewSession: () => new MediaTimingPreviewSession(),
    generateWaveform: (options) => generateSpeechWaveform(options),
    generateFrame: (options) => mediaTimingFrameExtractor.generate(options),
    clearFrameCache: () => mediaTimingFrameExtractor.clear(),
    resolveVideoSource: () =>
      resolveMediaGenerationInput(appState.mpvClient, 'video', getMediaGenerationOptions()),
    resolveMediaSource: async () => {
      const resolved = await resolveMediaGenerationInput(
        appState.mpvClient,
        'audio',
        getMediaGenerationOptions(),
      );
      return resolved
        ? {
            path: resolved.path,
            ...(resolved.inputOptions ? { inputOptions: resolved.inputOptions } : {}),
            singleResolvedStream: resolved.singleResolvedStream,
          }
        : null;
    },
    acquireMediaWindow: (source, range) => getSharedRemoteMediaWindowCache().acquire(source, range),
    getSubtitleContextLines: (range) =>
      collectMediaTimingContextLines({
        // Review ranges and tracker entries are on the video clock; parsed cues are not.
        cues: appState.activeParsedSubtitleCues.map((cue) =>
          shiftSubtitleRangeToVideoClock(cue, appState.mpvClient?.currentSubDelay ?? 0),
        ),
        fallbackPrevious: appState.subtitleTimingTracker?.getRecentEntries(40) ?? [],
        startTime: range.startTime,
        endTime: range.endTime,
      }),
    openModal: (payload, signal) =>
      openMediaTimingReviewModal(
        deps.getOverlayModals().createOverlayHostedModalOpenDeps(),
        payload,
        signal,
      ),
    onPreviewEnded: (reviewId) => {
      sendToOverlayWindows(IPC_CHANNELS.event.mediaTimingReviewPreviewEnded, reviewId);
    },
    onPreviewPosition: (position) => {
      sendToOverlayWindows(IPC_CHANNELS.event.mediaTimingReviewPreviewPosition, position);
    },
    showStatus: (message) =>
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification(message, { variant: 'warning' }),
  });

  const updateLastCardFromClipboard = createUpdateLastCardFromClipboardHandler({
    getAnkiIntegration: () => appState.ankiIntegration,
    readClipboardText: () => deps.clipboard.readText(),
    showMpvOsd: showStatus,
    updateLastCardFromClipboardCore,
  });

  const refreshKnownWordCache = createRefreshKnownWordCacheHandler({
    getAnkiIntegration: () => appState.ankiIntegration,
    missingIntegrationMessage: 'AnkiConnect integration not enabled',
  });

  const triggerFieldGrouping = createTriggerFieldGroupingHandler({
    getAnkiIntegration: () => appState.ankiIntegration,
    showMpvOsd: showStatus,
    triggerFieldGroupingCore,
  });

  const markLastCardAsAudioCard = createMarkLastCardAsAudioCardHandler({
    getAnkiIntegration: () => appState.ankiIntegration,
    showMpvOsd: showStatus,
    markLastCardAsAudioCardCore,
  });

  const mineSentenceCard = createMineSentenceCardHandler({
    getAnkiIntegration: () => appState.ankiIntegration,
    getMpvClient: () => appState.mpvClient,
    getPrimarySubtitle: () => captureCurrentPrimarySubtitleMiningContext(),
    showMpvOsd: showStatus,
    mineSentenceCardCore,
  });

  const copySubtitleCount = createHandleMultiCopyDigitHandler({
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    writeClipboardText: (text) => deps.clipboard.writeText(text),
    showMpvOsd: showPlaybackFeedback,
    handleMultiCopyDigitCore,
  });

  const copyCurrentSubtitle = createCopyCurrentSubtitleHandler({
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    writeClipboardText: (text) => deps.clipboard.writeText(text),
    showMpvOsd: showStatus,
    copyCurrentSubtitleCore,
  });

  const mineSentenceCount = createHandleMineSentenceDigitHandler({
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    getAnkiIntegration: () => appState.ankiIntegration,
    getCurrentSecondarySubText: () => appState.mpvClient?.currentSecondarySubText || undefined,
    showMpvOsd: showStatus,
    logError: (message, err) => {
      logger.error(message, err);
    },
    handleMineSentenceDigitCore,
  });

  /** Points a freshly created Anki integration at mining, stats, plugins and furigana. */
  function attachToAnkiIntegration(integration: AnkiIntegrationLike): void {
    const callbacks = deps.ankiIntegrationCallbacks;
    integration.setRecordCardsMinedCallback(callbacks.recordCardsMined);
    integration.setMiningProvenanceProvider(callbacks.getMiningProvenance);
    integration.setCardsReadyCallback(callbacks.onCardsReady);
    integration.setKnownWordCacheUpdatedCallback(callbacks.onKnownWordCacheUpdated);
    integration.setSubtitleMiningContextConsumer(consumePendingSubtitleMiningContext);
    integration.setMediaTimingReviewCallback(mediaTimingReviewRuntime.requestReview);
    integration.setSentenceFuriganaGenerator(callbacks.generateSentenceFurigana);
  }

  return {
    captureCurrentPrimarySubtitleMiningContext,
    recordSubtitleMiningContext,
    getFieldGroupingResolver,
    setFieldGroupingResolver,
    createFieldGroupingCallback: fieldGroupingOverlayRuntime.createFieldGroupingCallback,
    mediaTimingReviewRuntime,
    attachToAnkiIntegration,
    updateLastCardFromClipboard,
    refreshKnownWordCache,
    triggerFieldGrouping,
    markLastCardAsAudioCard,
    mineSentenceCard,
    mineSentenceCount,
    copyCurrentSubtitle,
    copySubtitleCount,
  };
}
