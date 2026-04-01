import type { SubtitleTimingTracker } from '../subtitle-timing-tracker';
import type { MpvSubtitleRenderMetrics } from '../types';
import type { MpvIpcClient } from '../core/services/mpv';
import { sendMpvCommandRuntime } from '../core/services';
import type { AnilistRuntime } from './anilist-runtime';
import type { DictionarySupportRuntime } from './dictionary-support-runtime';
import type { JellyfinRuntime } from './jellyfin-runtime';
import { createMiningRuntime } from './mining-runtime';
import type { MiningRuntimeInput } from './mining-runtime';
import { createMpvRuntimeFromMainState } from './mpv-runtime-bootstrap';
import type { MpvRuntime } from './mpv-runtime';
import type { SubtitleRuntime } from './subtitle-runtime';
import type { YoutubeRuntime } from './youtube-runtime';
import type { AppState } from './state';

export interface MainPlaybackRuntimeInput {
  appState: AppState;
  logPath: string;
  logger: Parameters<typeof createMpvRuntimeFromMainState>[0]['logger'] & {
    error: (message: string, error: unknown) => void;
  };
  getResolvedConfig: Parameters<typeof createMpvRuntimeFromMainState>[0]['getResolvedConfig'];
  getRuntimeBooleanOption: Parameters<
    typeof createMpvRuntimeFromMainState
  >[0]['getRuntimeBooleanOption'];
  subtitle: SubtitleRuntime;
  yomitan: {
    ensureYomitanExtensionLoaded: () => Promise<unknown>;
    isCharacterDictionaryEnabled: () => boolean;
  };
  currentMediaTokenizationGate: Parameters<
    typeof createMpvRuntimeFromMainState
  >[0]['currentMediaTokenizationGate'];
  startupOsdSequencer: Parameters<typeof createMpvRuntimeFromMainState>[0]['startupOsdSequencer'];
  dictionarySupport: DictionarySupportRuntime;
  overlay: {
    broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
    getVisibleOverlayVisible: () => boolean;
    getOverlayUi: () => { setOverlayVisible: (visible: boolean) => void } | undefined;
  };
  lifecycle: {
    requestAppQuit: () => void;
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
  mining: Omit<
    MiningRuntimeInput<any, any>,
    'showMpvOsd' | 'sendMpvCommand' | 'logError' | 'recordCardsMined'
  > & {
    readClipboardText: () => string;
    writeClipboardText: (text: string) => void;
    recordCardsMined: (count: number, noteIds?: number[]) => void;
  };
}

export interface MainPlaybackRuntime {
  mpvRuntime: MpvRuntime;
  mining: ReturnType<typeof createMiningRuntime>;
}

export function createMainPlaybackRuntime(input: MainPlaybackRuntimeInput): MainPlaybackRuntime {
  let mpvRuntime!: MpvRuntime;

  const showMpvOsd = (text: string): void => {
    mpvRuntime.showMpvOsd(text);
  };

  const mining = createMiningRuntime({
    ...input.mining,
    showMpvOsd: (text) => showMpvOsd(text),
    sendMpvCommand: (command) => {
      sendMpvCommandRuntime(input.appState.mpvClient, command);
    },
    logError: (message, err) => {
      input.logger.error(message, err);
    },
    recordCardsMined: (count, noteIds) => input.mining.recordCardsMined(count, noteIds),
  });

  mpvRuntime = createMpvRuntimeFromMainState({
    appState: input.appState,
    logPath: input.logPath,
    logger: input.logger,
    getResolvedConfig: input.getResolvedConfig,
    getRuntimeBooleanOption: input.getRuntimeBooleanOption,
    subtitle: input.subtitle,
    yomitan: {
      ensureYomitanExtensionLoaded: async () => {
        await input.yomitan.ensureYomitanExtensionLoaded();
      },
    },
    currentMediaTokenizationGate: input.currentMediaTokenizationGate,
    startupOsdSequencer: input.startupOsdSequencer,
    dictionarySupport: input.dictionarySupport,
    overlay: {
      broadcastToOverlayWindows: (channel, payload) => {
        input.overlay.broadcastToOverlayWindows(channel, payload);
      },
      getVisibleOverlayVisible: () => input.overlay.getVisibleOverlayVisible(),
      getOverlayUi: () => input.overlay.getOverlayUi(),
    },
    lifecycle: {
      requestAppQuit: () => input.lifecycle.requestAppQuit(),
      setQuitCheckTimer: (callback, timeoutMs) => {
        setTimeout(callback, timeoutMs);
      },
      restoreOverlayMpvSubtitles: input.lifecycle.restoreOverlayMpvSubtitles,
      syncOverlayMpvSubtitleSuppression: input.lifecycle.syncOverlayMpvSubtitleSuppression,
      publishDiscordPresence: () => input.lifecycle.publishDiscordPresence(),
    },
    stats: input.stats,
    anilist: input.anilist,
    jellyfin: input.jellyfin,
    youtube: input.youtube,
    isCharacterDictionaryEnabled: () => input.yomitan.isCharacterDictionaryEnabled(),
  }).mpvRuntime;

  return {
    mpvRuntime,
    mining,
  };
}
