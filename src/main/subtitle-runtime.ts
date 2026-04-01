import { createSubtitleProcessingController } from '../core/services/subtitle-processing-controller';
import {
  createSubtitlePrefetchService,
  type SubtitlePrefetchService,
  type SubtitlePrefetchServiceDeps,
} from '../core/services/subtitle-prefetch';
import type { SubtitleWebsocketFrequencyOptions } from '../core/services/subtitle-ws';
import type { SubtitleCue } from '../core/services/subtitle-cue-parser';
import {
  loadSubtitlePosition as loadSubtitlePositionCore,
  saveSubtitlePosition as saveSubtitlePositionCore,
} from '../core/services/subtitle-position';
import type {
  ResolvedConfig,
  SubtitleData,
  SubtitlePosition,
  SubtitleSidebarSnapshot,
} from '../types';
import {
  createLoadSubtitlePositionHandler,
  createSaveSubtitlePositionHandler,
} from './runtime/subtitle-position';
import { resolveSubtitleSourcePath } from './runtime/subtitle-prefetch-source';
import {
  createRefreshSubtitlePrefetchFromActiveTrackHandler,
  createResolveActiveSubtitleSidebarSourceHandler,
} from './runtime/subtitle-prefetch-runtime';
import { createSubtitlePrefetchInitController } from './runtime/subtitle-prefetch-init';
import {
  createExtractInternalSubtitleTrackToTempFileHandler,
  createSubtitleSourceLoader,
  type MpvSubtitleTrackLike,
} from './subtitle-runtime-sources';

type SubtitleBroadcastService = {
  broadcast: (payload: SubtitleData, options: SubtitleWebsocketFrequencyOptions) => void;
};

type SubtitleRuntimeConfigLike = Pick<
  ResolvedConfig,
  'subtitleStyle' | 'subtitleSidebar' | 'subtitlePosition' | 'subsync'
>;

export interface SubtitleRuntimeInput {
  getResolvedConfig: () => SubtitleRuntimeConfigLike;
  getCurrentMediaPath: () => string | null;
  getCurrentMediaTitle: () => string | null;
  getCurrentSubText: () => string;
  getCurrentSubAssText: () => string;
  getMpvClient: () => {
    connected?: boolean;
    currentSubStart?: number | null;
    currentSubEnd?: number | null;
    currentTimePos?: number | null;
    requestProperty: (name: string) => Promise<unknown>;
  } | null;
  subtitleWsService: SubtitleBroadcastService;
  annotationSubtitleWsService: SubtitleBroadcastService;
  broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  subtitlePositionsDir: string;
  setSubtitlePosition: (position: SubtitlePosition | null) => void;
  setPendingSubtitlePosition: (position: SubtitlePosition | null) => void;
  clearPendingSubtitlePosition: () => void;
  setCurrentSubtitleData?: (payload: SubtitleData | null) => void;
  setActiveParsedSubtitleState?: (cues: SubtitleCue[], sourceKey: string | null) => void;
  parseSubtitleCues: (content: string, filename: string) => SubtitleCue[];
  createSubtitlePrefetchService: (deps: SubtitlePrefetchServiceDeps) => SubtitlePrefetchService;
  loadSubtitleSourceText?: (source: string) => Promise<string>;
  refreshSubtitlePrefetchFromActiveTrack?: () => Promise<void>;
  fetchImpl?: typeof fetch;
  subtitleSourceFetchTimeoutMs?: number;
  prefetchRefreshDelayMs?: number;
  seekThresholdSeconds?: number;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  clearSchedule: (timer: ReturnType<typeof setTimeout>) => void;
  logDebug: (message: string) => void;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
}

export interface SubtitleRuntime {
  setTokenizeSubtitleDeferred: (tokenize: ((text: string) => Promise<SubtitleData>) | null) => void;
  emitSubtitlePayload: (payload: SubtitleData) => void;
  refreshCurrentSubtitle: (textOverride?: string) => void;
  invalidateTokenizationCache: () => void;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  consumeCachedSubtitle: (text: string) => SubtitleData | null;
  isCacheFull: () => boolean;
  onSubtitleChange: (text: string) => void;
  onCurrentMediaPathChange: (path: string | null) => void;
  onTimePosUpdate: (time: number) => void;
  getLastObservedTimePos: () => number;
  refreshSubtitleSidebarFromSource: (sourcePath: string) => Promise<void>;
  refreshSubtitlePrefetchFromActiveTrack: () => Promise<void>;
  cancelPendingSubtitlePrefetchInit: () => void;
  scheduleSubtitlePrefetchRefresh: (delayMs?: number) => void;
  clearScheduledSubtitlePrefetchRefresh: () => void;
  getSubtitleSidebarSnapshot: () => Promise<SubtitleSidebarSnapshot>;
  tokenizeCurrentSubtitle: () => Promise<SubtitleData>;
  loadSubtitleSourceText: (source: string) => Promise<string>;
  extractInternalSubtitleTrackToTempFile: (
    ffmpegPath: string,
    videoPath: string,
    track: MpvSubtitleTrackLike,
  ) => Promise<{ path: string; cleanup: () => Promise<void> } | null>;
  loadSubtitlePosition: () => SubtitlePosition | null;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getCurrentSubtitleData: () => SubtitleData | null;
  getActiveParsedSubtitleCues: () => SubtitleCue[];
  getActiveParsedSubtitleSource: () => string | null;
}

const DEFAULT_PREFETCH_REFRESH_DELAY_MS = 500;
const SEEK_THRESHOLD_SECONDS = 3;

export function createSubtitleRuntime(input: SubtitleRuntimeInput): SubtitleRuntime {
  const loadSubtitleSourceText =
    input.loadSubtitleSourceText ??
    createSubtitleSourceLoader({
      fetchImpl: input.fetchImpl,
      subtitleSourceFetchTimeoutMs: input.subtitleSourceFetchTimeoutMs,
    });
  const extractInternalSubtitleTrackToTempFile =
    createExtractInternalSubtitleTrackToTempFileHandler();

  let tokenizeSubtitleDeferred: ((text: string) => Promise<SubtitleData>) | null = null;
  let currentSubtitleData: SubtitleData | null = null;
  let subtitlePrefetchService: SubtitlePrefetchService | null = null;
  let subtitlePrefetchRefreshTimer: ReturnType<typeof setTimeout> | null = null;
  let lastObservedTimePos = 0;
  let activeParsedSubtitleCues: SubtitleCue[] = [];
  let activeParsedSubtitleSource: string | null = null;

  const setActiveParsedSubtitleState = (
    cues: SubtitleCue[] | null,
    sourceKey: string | null,
  ): void => {
    activeParsedSubtitleCues = cues ?? [];
    activeParsedSubtitleSource = sourceKey;
    input.setActiveParsedSubtitleState?.(activeParsedSubtitleCues, activeParsedSubtitleSource);
  };

  const tokenizationController = createSubtitleProcessingController({
    tokenizeSubtitle: async (text: string) =>
      tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : { text, tokens: null },
    emitSubtitle: (payload) => emitSubtitlePayload(payload),
    logDebug: (message) => {
      input.logDebug(`[subtitle-processing] ${message}`);
    },
    now: () => Date.now(),
  });

  const subtitlePrefetchInitController = createSubtitlePrefetchInitController({
    getCurrentService: () => subtitlePrefetchService,
    setCurrentService: (service) => {
      subtitlePrefetchService = service;
    },
    loadSubtitleSourceText,
    parseSubtitleCues: (content, filename) => input.parseSubtitleCues(content, filename),
    createSubtitlePrefetchService: (deps) => input.createSubtitlePrefetchService(deps),
    tokenizeSubtitle: async (text) =>
      tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
    preCacheTokenization: (text, data) => tokenizationController.preCacheTokenization(text, data),
    isCacheFull: () => tokenizationController.isCacheFull(),
    logInfo: (message) => input.logInfo(message),
    logWarn: (message) => input.logWarn(message),
    onParsedSubtitleCuesChanged: (cues, sourceKey) => {
      setActiveParsedSubtitleState(cues, sourceKey);
    },
  });

  const resolveActiveSubtitleSidebarSourceHandler = createResolveActiveSubtitleSidebarSourceHandler(
    {
      getFfmpegPath: () => input.getResolvedConfig().subsync.ffmpeg_path.trim() || 'ffmpeg',
      extractInternalSubtitleTrack: (ffmpegPath, videoPath, track) =>
        extractInternalSubtitleTrackToTempFile(ffmpegPath, videoPath, track),
    },
  );

  const refreshSubtitlePrefetchFromActiveTrackHandler =
    input.refreshSubtitlePrefetchFromActiveTrack ??
    createRefreshSubtitlePrefetchFromActiveTrackHandler({
      getMpvClient: () => input.getMpvClient(),
      getLastObservedTimePos: () => lastObservedTimePos,
      subtitlePrefetchInitController,
      resolveActiveSubtitleSidebarSource: (nextInput) =>
        resolveActiveSubtitleSidebarSourceHandler(nextInput),
    });

  const loadSubtitlePosition = createLoadSubtitlePositionHandler({
    loadSubtitlePositionCore: () =>
      loadSubtitlePositionCore({
        currentMediaPath: input.getCurrentMediaPath(),
        fallbackPosition: input.getResolvedConfig().subtitlePosition,
        subtitlePositionsDir: input.subtitlePositionsDir,
      }),
    setSubtitlePosition: (position) => input.setSubtitlePosition(position),
  });

  const saveSubtitlePosition = createSaveSubtitlePositionHandler({
    saveSubtitlePositionCore: (position) =>
      saveSubtitlePositionCore({
        position,
        currentMediaPath: input.getCurrentMediaPath(),
        subtitlePositionsDir: input.subtitlePositionsDir,
        onQueuePending: (queued) => input.setPendingSubtitlePosition(queued),
        onPersisted: () => input.clearPendingSubtitlePosition(),
      }),
    setSubtitlePosition: (position) => input.setSubtitlePosition(position),
  });

  const getSubtitleBroadcastOptions = (): SubtitleWebsocketFrequencyOptions => {
    const config = input.getResolvedConfig().subtitleStyle.frequencyDictionary;
    return {
      enabled: config.enabled,
      topX: config.topX,
      mode: config.mode,
    };
  };

  const withCurrentSubtitleTiming = (payload: SubtitleData): SubtitleData => ({
    ...payload,
    startTime: input.getMpvClient()?.currentSubStart ?? null,
    endTime: input.getMpvClient()?.currentSubEnd ?? null,
  });

  const clearScheduledSubtitlePrefetchRefresh = (): void => {
    if (subtitlePrefetchRefreshTimer) {
      input.clearSchedule(subtitlePrefetchRefreshTimer);
      subtitlePrefetchRefreshTimer = null;
    }
  };

  const scheduleSubtitlePrefetchRefresh = (delayMs = 0): void => {
    clearScheduledSubtitlePrefetchRefresh();
    subtitlePrefetchRefreshTimer = input.schedule(() => {
      subtitlePrefetchRefreshTimer = null;
      void refreshSubtitlePrefetchFromActiveTrackHandler();
    }, delayMs);
  };

  const refreshSubtitleSidebarFromSource = async (sourcePath: string): Promise<void> => {
    const normalizedSourcePath = resolveSubtitleSourcePath(sourcePath.trim());
    if (!normalizedSourcePath) {
      return;
    }
    await subtitlePrefetchInitController.initSubtitlePrefetch(
      normalizedSourcePath,
      lastObservedTimePos,
      normalizedSourcePath,
    );
  };

  const onCurrentMediaPathChange = (pathValue: string | null): void => {
    clearScheduledSubtitlePrefetchRefresh();
    subtitlePrefetchInitController.cancelPendingInit();
    if (pathValue) {
      scheduleSubtitlePrefetchRefresh(
        input.prefetchRefreshDelayMs ?? DEFAULT_PREFETCH_REFRESH_DELAY_MS,
      );
    }
  };

  const onTimePosUpdate = (time: number): void => {
    const delta = time - lastObservedTimePos;
    if (
      subtitlePrefetchService &&
      (delta > (input.seekThresholdSeconds ?? SEEK_THRESHOLD_SECONDS) || delta < 0)
    ) {
      subtitlePrefetchService.onSeek(time);
    }
    lastObservedTimePos = time;
  };

  const emitSubtitlePayload = (payload: SubtitleData): void => {
    const timedPayload = withCurrentSubtitleTiming(payload);
    currentSubtitleData = timedPayload;
    input.setCurrentSubtitleData?.(timedPayload);
    input.broadcastToOverlayWindows('subtitle:set', timedPayload);
    const broadcastOptions = getSubtitleBroadcastOptions();
    input.subtitleWsService.broadcast(timedPayload, broadcastOptions);
    input.annotationSubtitleWsService.broadcast(timedPayload, broadcastOptions);
    subtitlePrefetchService?.resume();
  };

  const tokenizeCurrentSubtitle = async (): Promise<SubtitleData> => {
    const tokenized = await tokenizationController.consumeCachedSubtitle(input.getCurrentSubText());
    if (tokenized) {
      return withCurrentSubtitleTiming(tokenized);
    }
    const text = input.getCurrentSubText();
    const deferred = tokenizeSubtitleDeferred
      ? await tokenizeSubtitleDeferred(text)
      : { text, tokens: null };
    return withCurrentSubtitleTiming(deferred);
  };

  return {
    setTokenizeSubtitleDeferred: (tokenize) => {
      tokenizeSubtitleDeferred = tokenize;
    },
    emitSubtitlePayload,
    refreshCurrentSubtitle: (textOverride?: string) => {
      tokenizationController.refreshCurrentSubtitle(textOverride);
    },
    invalidateTokenizationCache: () => {
      tokenizationController.invalidateTokenizationCache();
    },
    preCacheTokenization: (text, data) => {
      tokenizationController.preCacheTokenization(text, data);
    },
    consumeCachedSubtitle: (text) => tokenizationController.consumeCachedSubtitle(text),
    isCacheFull: () => tokenizationController.isCacheFull(),
    onSubtitleChange: (text) => {
      subtitlePrefetchService?.pause();
      tokenizationController.onSubtitleChange(text);
    },
    onCurrentMediaPathChange,
    onTimePosUpdate,
    getLastObservedTimePos: () => lastObservedTimePos,
    refreshSubtitleSidebarFromSource,
    refreshSubtitlePrefetchFromActiveTrack: async () => {
      await refreshSubtitlePrefetchFromActiveTrackHandler();
    },
    cancelPendingSubtitlePrefetchInit: () => {
      subtitlePrefetchInitController.cancelPendingInit();
    },
    scheduleSubtitlePrefetchRefresh,
    clearScheduledSubtitlePrefetchRefresh,
    getSubtitleSidebarSnapshot: async (): Promise<SubtitleSidebarSnapshot> => {
      const currentSubtitle = {
        text: input.getCurrentSubText(),
        startTime: input.getMpvClient()?.currentSubStart ?? null,
        endTime: input.getMpvClient()?.currentSubEnd ?? null,
      };
      const currentTimeSec = input.getMpvClient()?.currentTimePos ?? null;
      const config = input.getResolvedConfig().subtitleSidebar;
      const client = input.getMpvClient();
      if (!client?.connected) {
        return {
          cues: activeParsedSubtitleCues,
          currentTimeSec,
          currentSubtitle,
          config,
        };
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
          return {
            cues: activeParsedSubtitleCues,
            currentTimeSec,
            currentSubtitle,
            config,
          };
        }

        const resolvedSource = await resolveActiveSubtitleSidebarSourceHandler({
          currentExternalFilenameRaw,
          currentTrackRaw,
          trackListRaw,
          sidRaw,
          videoPath,
        });
        if (!resolvedSource) {
          return {
            cues: activeParsedSubtitleCues,
            currentTimeSec,
            currentSubtitle,
            config,
          };
        }

        try {
          if (activeParsedSubtitleSource === resolvedSource.sourceKey) {
            return {
              cues: activeParsedSubtitleCues,
              currentTimeSec,
              currentSubtitle,
              config,
            };
          }

          const content = await loadSubtitleSourceText(resolvedSource.path);
          const cues = input.parseSubtitleCues(content, resolvedSource.path);
          setActiveParsedSubtitleState(cues, resolvedSource.sourceKey);
          return {
            cues,
            currentTimeSec,
            currentSubtitle,
            config,
          };
        } finally {
          await resolvedSource.cleanup?.();
        }
      } catch {
        return {
          cues: activeParsedSubtitleCues,
          currentTimeSec,
          currentSubtitle,
          config,
        };
      }
    },
    tokenizeCurrentSubtitle,
    loadSubtitleSourceText,
    extractInternalSubtitleTrackToTempFile,
    loadSubtitlePosition,
    saveSubtitlePosition,
    getCurrentSubtitleData: () => currentSubtitleData,
    getActiveParsedSubtitleCues: () => activeParsedSubtitleCues,
    getActiveParsedSubtitleSource: () => activeParsedSubtitleSource,
  };
}
