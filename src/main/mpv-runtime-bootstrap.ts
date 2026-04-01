import type { MpvRuntime, MpvRuntimeInput } from './mpv-runtime';
import { createMpvRuntime } from './mpv-runtime';
import type { SubtitleRuntime } from './subtitle-runtime';
import type { DictionarySupportRuntime } from './dictionary-support-runtime';
import type { createCurrentMediaTokenizationGate } from './runtime/current-media-tokenization-gate';
import type { createStartupOsdSequencer } from './runtime/startup-osd-sequencer';
import type { AnilistRuntime } from './anilist-runtime';
import type { JellyfinRuntime } from './jellyfin-runtime';
import type { YoutubeRuntime } from './youtube-runtime';

export interface MpvRuntimeBootstrapInput {
  appState: MpvRuntimeInput['appState'];
  logPath: string;
  logger: MpvRuntimeInput['logger'];
  getResolvedConfig: MpvRuntimeInput['getResolvedConfig'];
  getRuntimeBooleanOption: MpvRuntimeInput['getRuntimeBooleanOption'];
  subtitle: MpvRuntimeInput['subtitle'];
  ensureYomitanExtensionLoaded: MpvRuntimeInput['ensureYomitanExtensionLoaded'];
  currentMediaTokenizationGate: MpvRuntimeInput['currentMediaTokenizationGate'];
  startupOsdSequencer: MpvRuntimeInput['startupOsdSequencer'];
  dictionarySupport: {
    ensureJlptDictionaryLookup: () => Promise<void>;
    ensureFrequencyDictionaryLookup: () => Promise<void>;
    syncImmersionMediaState: () => void;
    updateCurrentMediaPath: (mediaPath: unknown) => void;
    updateCurrentMediaTitle: (mediaTitle: unknown) => void;
    scheduleCharacterDictionarySync: () => void;
  };
  overlay: {
    overlayManager: {
      broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
      getVisibleOverlayVisible: () => boolean;
    };
    getOverlayUi: () => { setOverlayVisible: (visible: boolean) => void } | undefined;
  };
  lifecycle: {
    requestAppQuit: () => void;
    setQuitCheckTimer: (callback: () => void, timeoutMs: number) => void;
    restoreOverlayMpvSubtitles: () => void;
    syncOverlayMpvSubtitleSuppression: () => void;
    publishDiscordPresence: () => void;
  };
  stats: MpvRuntimeInput['stats'];
  anilist: MpvRuntimeInput['anilist'];
  jellyfin: MpvRuntimeInput['jellyfin'];
  youtube: MpvRuntimeInput['youtube'];
  isCharacterDictionaryEnabled: MpvRuntimeInput['isCharacterDictionaryEnabled'];
}

export interface MpvRuntimeBootstrap {
  mpvRuntime: MpvRuntime;
}

export interface MpvRuntimeFromMainStateInput {
  appState: MpvRuntimeInput['appState'];
  logPath: string;
  logger: MpvRuntimeInput['logger'];
  getResolvedConfig: MpvRuntimeInput['getResolvedConfig'];
  getRuntimeBooleanOption: MpvRuntimeInput['getRuntimeBooleanOption'];
  subtitle: SubtitleRuntime;
  yomitan: {
    ensureYomitanExtensionLoaded: () => Promise<void>;
  };
  currentMediaTokenizationGate: ReturnType<typeof createCurrentMediaTokenizationGate>;
  startupOsdSequencer: ReturnType<typeof createStartupOsdSequencer>;
  dictionarySupport: DictionarySupportRuntime;
  overlay: {
    broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
    getVisibleOverlayVisible: () => boolean;
    getOverlayUi: () => { setOverlayVisible: (visible: boolean) => void } | undefined;
  };
  lifecycle: {
    requestAppQuit: () => void;
    setQuitCheckTimer: (callback: () => void, timeoutMs: number) => void;
    restoreOverlayMpvSubtitles: () => void;
    syncOverlayMpvSubtitleSuppression: () => void;
    publishDiscordPresence: () => void;
  };
  stats: {
    ensureImmersionTrackerStarted: () => void;
  };
  anilist: AnilistRuntime;
  jellyfin: JellyfinRuntime;
  youtube: YoutubeRuntime;
  isCharacterDictionaryEnabled: () => boolean;
}

export function createMpvRuntimeBootstrap(input: MpvRuntimeBootstrapInput): MpvRuntimeBootstrap {
  const mpvRuntime = createMpvRuntime({
    appState: input.appState,
    logPath: input.logPath,
    logger: input.logger,
    getResolvedConfig: input.getResolvedConfig,
    getRuntimeBooleanOption: input.getRuntimeBooleanOption,
    subtitle: input.subtitle,
    ensureYomitanExtensionLoaded: input.ensureYomitanExtensionLoaded,
    currentMediaTokenizationGate: input.currentMediaTokenizationGate,
    startupOsdSequencer: input.startupOsdSequencer,
    dictionaries: {
      ensureJlptDictionaryLookup: () => input.dictionarySupport.ensureJlptDictionaryLookup(),
      ensureFrequencyDictionaryLookup: () =>
        input.dictionarySupport.ensureFrequencyDictionaryLookup(),
    },
    mediaRuntime: {
      syncImmersionMediaState: () => input.dictionarySupport.syncImmersionMediaState(),
      updateCurrentMediaPath: (mediaPath) => {
        input.dictionarySupport.updateCurrentMediaPath(mediaPath);
      },
      updateCurrentMediaTitle: (mediaTitle) => {
        input.dictionarySupport.updateCurrentMediaTitle(mediaTitle);
      },
    },
    characterDictionaryAutoSyncRuntime: {
      scheduleSync: () => {
        input.dictionarySupport.scheduleCharacterDictionarySync();
      },
    },
    overlay: {
      broadcastToOverlayWindows: (channel, payload) => {
        input.overlay.overlayManager.broadcastToOverlayWindows(channel, payload);
      },
      getVisibleOverlayVisible: () => input.overlay.overlayManager.getVisibleOverlayVisible(),
      setOverlayVisible: (visible) => {
        input.overlay.getOverlayUi()?.setOverlayVisible(visible);
      },
    },
    lifecycle: {
      requestAppQuit: () => input.lifecycle.requestAppQuit(),
      scheduleQuitCheck: (callback) => {
        input.lifecycle.setQuitCheckTimer(callback, 500);
      },
      restoreOverlayMpvSubtitles: () => {
        input.lifecycle.restoreOverlayMpvSubtitles();
      },
      syncOverlayMpvSubtitleSuppression: () => {
        input.lifecycle.syncOverlayMpvSubtitleSuppression();
      },
      refreshDiscordPresence: () => {
        input.lifecycle.publishDiscordPresence();
      },
    },
    stats: input.stats,
    anilist: input.anilist,
    jellyfin: input.jellyfin,
    youtube: input.youtube,
    isCharacterDictionaryEnabled: input.isCharacterDictionaryEnabled,
  });

  return {
    mpvRuntime,
  };
}

export function createMpvRuntimeFromMainState(
  input: MpvRuntimeFromMainStateInput,
): MpvRuntimeBootstrap {
  return createMpvRuntimeBootstrap({
    appState: input.appState,
    logPath: input.logPath,
    logger: input.logger,
    getResolvedConfig: input.getResolvedConfig,
    getRuntimeBooleanOption: input.getRuntimeBooleanOption,
    subtitle: {
      consumeCachedSubtitle: (text) => input.subtitle.consumeCachedSubtitle(text),
      emitSubtitlePayload: (payload) => input.subtitle.emitSubtitlePayload(payload),
      onSubtitleChange: (text) => {
        input.subtitle.onSubtitleChange(text);
      },
      onCurrentMediaPathChange: (path) => {
        input.subtitle.onCurrentMediaPathChange(path);
      },
      onTimePosUpdate: (time) => {
        input.subtitle.onTimePosUpdate(time);
      },
      scheduleSubtitlePrefetchRefresh: (delayMs) =>
        input.subtitle.scheduleSubtitlePrefetchRefresh(delayMs),
      loadSubtitleSourceText: (source) => input.subtitle.loadSubtitleSourceText(source),
      setTokenizeSubtitleDeferred: (tokenize) => {
        input.subtitle.setTokenizeSubtitleDeferred(tokenize);
      },
    },
    ensureYomitanExtensionLoaded: async () => {
      await input.yomitan.ensureYomitanExtensionLoaded();
    },
    currentMediaTokenizationGate: input.currentMediaTokenizationGate,
    startupOsdSequencer: input.startupOsdSequencer,
    dictionarySupport: {
      ensureJlptDictionaryLookup: () => input.dictionarySupport.ensureJlptDictionaryLookup(),
      ensureFrequencyDictionaryLookup: () =>
        input.dictionarySupport.ensureFrequencyDictionaryLookup(),
      syncImmersionMediaState: () => {
        input.dictionarySupport.syncImmersionMediaState();
      },
      updateCurrentMediaPath: (mediaPath) => {
        input.dictionarySupport.updateCurrentMediaPath(mediaPath);
      },
      updateCurrentMediaTitle: (mediaTitle) => {
        input.dictionarySupport.updateCurrentMediaTitle(mediaTitle);
      },
      scheduleCharacterDictionarySync: () => {
        input.dictionarySupport.scheduleCharacterDictionarySync();
      },
    },
    overlay: {
      overlayManager: {
        broadcastToOverlayWindows: (channel, payload) => {
          input.overlay.broadcastToOverlayWindows(channel, payload);
        },
        getVisibleOverlayVisible: () => input.overlay.getVisibleOverlayVisible(),
      },
      getOverlayUi: () => input.overlay.getOverlayUi(),
    },
    lifecycle: {
      requestAppQuit: () => input.lifecycle.requestAppQuit(),
      setQuitCheckTimer: (callback, timeoutMs) => {
        input.lifecycle.setQuitCheckTimer(callback, timeoutMs);
      },
      restoreOverlayMpvSubtitles: () => {
        input.lifecycle.restoreOverlayMpvSubtitles();
      },
      syncOverlayMpvSubtitleSuppression: () => {
        input.lifecycle.syncOverlayMpvSubtitleSuppression();
      },
      publishDiscordPresence: () => {
        input.lifecycle.publishDiscordPresence();
      },
    },
    stats: {
      ensureImmersionTrackerStarted: () => input.stats.ensureImmersionTrackerStarted(),
    },
    anilist: {
      getCurrentAnilistMediaKey: () => input.anilist.getCurrentAnilistMediaKey(),
      resetAnilistMediaTracking: (mediaKey) => {
        input.anilist.resetAnilistMediaTracking(mediaKey);
      },
      maybeProbeAnilistDuration: (mediaKey) => {
        if (mediaKey) {
          void input.anilist.maybeProbeAnilistDuration(mediaKey);
        }
      },
      ensureAnilistMediaGuess: (mediaKey) => {
        if (mediaKey) {
          void input.anilist.ensureAnilistMediaGuess(mediaKey);
        }
      },
      maybeRunAnilistPostWatchUpdate: () => input.anilist.maybeRunAnilistPostWatchUpdate(),
      resetAnilistMediaGuessState: () => {
        input.anilist.resetAnilistMediaGuessState();
      },
    },
    jellyfin: {
      getQuitOnDisconnectArmed: () => input.jellyfin.getQuitOnDisconnectArmed(),
      reportJellyfinRemoteStopped: () => input.jellyfin.reportJellyfinRemoteStopped(),
      reportJellyfinRemoteProgress: (forceImmediate) =>
        input.jellyfin.reportJellyfinRemoteProgress(forceImmediate),
      startJellyfinRemoteSession: () => input.jellyfin.startJellyfinRemoteSession(),
    },
    youtube: {
      getQuitOnDisconnectArmed: () => input.youtube.getQuitOnDisconnectArmed(),
      handleMpvConnectionChange: (connected) => {
        input.youtube.handleMpvConnectionChange(connected);
      },
      handleMediaPathChange: (path) => {
        input.youtube.invalidatePendingAutoplayReadyFallbacks();
        input.currentMediaTokenizationGate.updateCurrentMediaPath(path);
        input.startupOsdSequencer.reset();
        input.youtube.handleMediaPathChange(path);
        if (path) {
          input.stats.ensureImmersionTrackerStarted();
        }
      },
      handleSubtitleTrackChange: (sid) => {
        input.youtube.handleSubtitleTrackChange(sid);
      },
      handleSubtitleTrackListChange: (trackList) => {
        input.youtube.handleSubtitleTrackListChange(trackList);
      },
      invalidatePendingAutoplayReadyFallbacks: () =>
        input.youtube.invalidatePendingAutoplayReadyFallbacks(),
      maybeSignalPluginAutoplayReady: (subtitle, options) =>
        input.youtube.maybeSignalPluginAutoplayReady(subtitle, options),
    },
    isCharacterDictionaryEnabled: input.isCharacterDictionaryEnabled,
  });
}
