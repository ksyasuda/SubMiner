import type { CliArgs, CliCommandSource } from '../cli/args';
import type {
  SubtitleData,
  YoutubePickerOpenPayload,
  YoutubePickerResolveRequest,
  YoutubePickerResolveResult,
} from '../types';
import type {
  YoutubeTrackOption,
  YoutubeTrackProbeResult,
} from '../core/services/youtube/track-probe';
import { createAutoplayReadyGate } from './runtime/autoplay-ready-gate';
import { createYoutubeFlowRuntime } from './runtime/youtube-flow';
import { createYoutubePlaybackRuntime } from './runtime/youtube-playback-runtime';
import {
  clearYoutubePrimarySubtitleNotificationTimer,
  createYoutubePrimarySubtitleNotificationRuntime,
} from './runtime/youtube-primary-subtitle-notification';
import { isYoutubePlaybackActive } from './runtime/youtube-playback';

type YoutubeFlowRuntimeLike = {
  runYoutubePlaybackFlow: (request: {
    url: string;
    mode: 'download' | 'generate';
  }) => Promise<void>;
  openManualPicker: (request: { url: string }) => Promise<void>;
  resolveActivePicker: (
    request: YoutubePickerResolveRequest,
  ) => Promise<YoutubePickerResolveResult>;
  cancelActivePicker: () => boolean;
  hasActiveSession: () => boolean;
};

type YoutubePlaybackRuntimeLike = {
  clearYoutubePlayQuitOnDisconnectArmTimer: () => void;
  getQuitOnDisconnectArmed: () => boolean;
  runYoutubePlaybackFlow: (request: {
    url: string;
    mode: NonNullable<CliArgs['youtubeMode']>;
    source: CliCommandSource;
  }) => Promise<void>;
};

type YoutubeAutoplayGateLike = {
  getAutoPlayReadySignalMediaPath: () => string | null;
  invalidatePendingAutoplayReadyFallbacks: () => void;
  maybeSignalPluginAutoplayReady: (
    payload: SubtitleData,
    options?: { forceWhilePaused?: boolean },
  ) => void;
};

type YoutubePrimarySubtitleNotificationRuntimeLike = {
  handleMediaPathChange: (path: string | null) => void;
  handleSubtitleTrackChange: (sid: number | null) => void;
  handleSubtitleTrackListChange: (trackList: unknown[] | null) => void;
  setAppOwnedFlowInFlight: (inFlight: boolean) => void;
  isAppOwnedFlowInFlight: () => boolean;
};

export interface YoutubeFlowRuntimeInput {
  probeYoutubeTracks: (url: string) => Promise<YoutubeTrackProbeResult>;
  acquireYoutubeSubtitleTrack: (input: {
    targetUrl: string;
    outputDir: string;
    track: YoutubeTrackOption;
  }) => Promise<{ path: string }>;
  acquireYoutubeSubtitleTracks: (input: {
    targetUrl: string;
    outputDir: string;
    tracks: YoutubeTrackOption[];
  }) => Promise<Map<string, string>>;
  openPicker: (payload: YoutubePickerOpenPayload) => Promise<boolean>;
  pauseMpv: () => void;
  resumeMpv: () => void;
  sendMpvCommand: (command: Array<string | number>) => void;
  requestMpvProperty: (name: string) => Promise<unknown>;
  refreshCurrentSubtitle: (text: string) => void;
  refreshSubtitleSidebarSource?: (sourcePath: string) => Promise<void>;
  startTokenizationWarmups: () => Promise<void>;
  waitForTokenizationReady: () => Promise<void>;
  waitForAnkiReady: () => Promise<void>;
  wait: (ms: number) => Promise<void>;
  waitForPlaybackWindowReady: () => Promise<void>;
  waitForOverlayGeometryReady: () => Promise<void>;
  focusOverlayWindow: () => void;
  showMpvOsd: (text: string) => void;
  warn: (message: string) => void;
  log: (message: string) => void;
  getYoutubeOutputDir: () => string;
}

export interface YoutubePlaybackRuntimeInput {
  platform: NodeJS.Platform;
  directPlaybackFormat: string;
  mpvYtdlFormat: string;
  autoLaunchTimeoutMs: number;
  connectTimeoutMs: number;
  getSocketPath: () => string;
  getMpvConnected: () => boolean;
  ensureYoutubePlaybackRuntimeReady: () => Promise<void>;
  resolveYoutubePlaybackUrl: (url: string, format: string) => Promise<string>;
  launchWindowsMpv: (
    playbackUrl: string,
    args: string[],
  ) => {
    ok: boolean;
    mpvPath?: string;
  };
  waitForYoutubeMpvConnected: (timeoutMs: number) => Promise<boolean>;
  prepareYoutubePlaybackInMpv: (request: { url: string }) => Promise<boolean>;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  clearScheduled: (timer: ReturnType<typeof setTimeout>) => void;
}

export interface YoutubeAutoplayRuntimeInput {
  getCurrentMediaPath: () => string | null;
  getCurrentVideoPath: () => string | null;
  getPlaybackPaused: () => boolean | null;
  getMpvClient: () => {
    connected?: boolean;
    requestProperty: (property: string) => Promise<unknown>;
    send: (payload: { command: Array<string | boolean> }) => void;
  } | null;
  signalPluginAutoplayReady: () => void;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  logDebug: (message: string) => void;
}

export interface YoutubeNotificationRuntimeInput {
  getPrimarySubtitleLanguages: () => string[];
  schedule: (
    callback: () => void,
    delayMs: number,
  ) => ReturnType<typeof setTimeout> | { id: number };
  clearSchedule: (timer: ReturnType<typeof setTimeout> | { id: number } | null) => void;
}

export interface YoutubeRuntimeInput {
  flow: YoutubeFlowRuntimeInput;
  playback: YoutubePlaybackRuntimeInput;
  autoplay: YoutubeAutoplayRuntimeInput;
  notification: YoutubeNotificationRuntimeInput;
  getNotificationType: () => 'osd' | 'system' | 'both' | string;
  getCurrentMediaPath: () => string | null;
  getCurrentVideoPath: () => string | null;
  showMpvOsd: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  broadcastYoutubePickerCancel: () => void;
  closeYoutubePickerModal: () => void;
  logWarn: (message: string) => void;
  createFlowRuntime?: (
    input: YoutubeFlowRuntimeInput & { reportSubtitleFailure: (message: string) => void },
  ) => YoutubeFlowRuntimeLike;
  createPlaybackRuntime?: (
    input: YoutubePlaybackRuntimeInput & {
      invalidatePendingAutoplayReadyFallbacks: () => void;
      setAppOwnedFlowInFlight: (next: boolean) => void;
      runYoutubePlaybackFlow: (request: {
        url: string;
        mode: NonNullable<CliArgs['youtubeMode']>;
      }) => Promise<void>;
    },
  ) => YoutubePlaybackRuntimeLike;
  createAutoplayGate?: (
    input: {
      isAppOwnedFlowInFlight: () => boolean;
    } & YoutubeAutoplayRuntimeInput,
  ) => YoutubeAutoplayGateLike;
  createPrimarySubtitleNotificationRuntime?: (
    input: {
      notifyFailure: (message: string) => void;
    } & YoutubeNotificationRuntimeInput,
  ) => YoutubePrimarySubtitleNotificationRuntimeLike;
}

export interface YoutubeRuntime {
  runYoutubePlaybackFlow: (request: {
    url: string;
    mode: NonNullable<CliArgs['youtubeMode']>;
    source: CliCommandSource;
  }) => Promise<void>;
  openYoutubeTrackPickerFromPlayback: () => Promise<void>;
  resolveActivePicker: (
    request: YoutubePickerResolveRequest,
  ) => Promise<YoutubePickerResolveResult>;
  handleMpvConnectionChange: (connected: boolean) => void;
  reportYoutubeSubtitleFailure: (message: string) => void;
  handleMediaPathChange: (path: string | null) => void;
  handleSubtitleTrackChange: (sid: number | null) => void;
  handleSubtitleTrackListChange: (trackList: unknown[] | null) => void;
  maybeSignalPluginAutoplayReady: (
    payload: SubtitleData,
    options?: { forceWhilePaused?: boolean },
  ) => void;
  invalidatePendingAutoplayReadyFallbacks: () => void;
  getAutoPlayReadySignalMediaPath: () => string | null;
  clearYoutubePlayQuitOnDisconnectArmTimer: () => void;
  getQuitOnDisconnectArmed: () => boolean;
  isAppOwnedFlowInFlight: () => boolean;
}

export function createYoutubeRuntime(input: YoutubeRuntimeInput): YoutubeRuntime {
  const reportYoutubeSubtitleFailure = (message: string): void => {
    const type = input.getNotificationType();
    if (type === 'osd' || type === 'both') {
      input.showMpvOsd(message);
    }
    if (type === 'system' || type === 'both') {
      try {
        input.showDesktopNotification('SubMiner', { body: message });
      } catch {
        input.logWarn(`Unable to show desktop notification: ${message}`);
      }
    }
  };

  const notificationRuntime = (
    input.createPrimarySubtitleNotificationRuntime ??
    ((deps) => createYoutubePrimarySubtitleNotificationRuntime(deps))
  )({
    ...input.notification,
    notifyFailure: (message) => reportYoutubeSubtitleFailure(message),
  });

  const autoplayGate = (input.createAutoplayGate ?? ((deps) => createAutoplayReadyGate(deps)))({
    ...input.autoplay,
    isAppOwnedFlowInFlight: () => notificationRuntime.isAppOwnedFlowInFlight(),
  });

  const flowRuntime = (
    input.createFlowRuntime ?? ((deps) => createYoutubeFlowRuntime(deps as never))
  )({
    ...input.flow,
    reportSubtitleFailure: (message) => reportYoutubeSubtitleFailure(message),
  });

  const playbackRuntime = (
    input.createPlaybackRuntime ?? ((deps) => createYoutubePlaybackRuntime(deps))
  )({
    ...input.playback,
    invalidatePendingAutoplayReadyFallbacks: () =>
      autoplayGate.invalidatePendingAutoplayReadyFallbacks(),
    setAppOwnedFlowInFlight: (next) => {
      notificationRuntime.setAppOwnedFlowInFlight(next);
    },
    runYoutubePlaybackFlow: (request) => flowRuntime.runYoutubePlaybackFlow(request),
  });

  const isYoutubePlaybackActiveNow = (): boolean =>
    isYoutubePlaybackActive(input.getCurrentMediaPath(), input.getCurrentVideoPath());

  const openYoutubeTrackPickerFromPlayback = async (): Promise<void> => {
    if (flowRuntime.hasActiveSession()) {
      input.showMpvOsd('YouTube subtitle flow already in progress.');
      return;
    }

    const currentMediaPath =
      input.getCurrentMediaPath()?.trim() || input.getCurrentVideoPath()?.trim() || '';
    if (!isYoutubePlaybackActiveNow() || !currentMediaPath) {
      input.showMpvOsd('YouTube subtitle picker is only available during YouTube playback.');
      return;
    }

    await flowRuntime.openManualPicker({
      url: currentMediaPath,
    });
  };

  const handleMpvConnectionChange = (connected: boolean): void => {
    if (connected || !flowRuntime.hasActiveSession()) {
      return;
    }
    flowRuntime.cancelActivePicker();
    input.broadcastYoutubePickerCancel();
    input.closeYoutubePickerModal();
  };

  return {
    runYoutubePlaybackFlow: (request) => playbackRuntime.runYoutubePlaybackFlow(request),
    openYoutubeTrackPickerFromPlayback,
    resolveActivePicker: (request) => flowRuntime.resolveActivePicker(request),
    handleMpvConnectionChange,
    reportYoutubeSubtitleFailure,
    handleMediaPathChange: (path) => notificationRuntime.handleMediaPathChange(path),
    handleSubtitleTrackChange: (sid) => notificationRuntime.handleSubtitleTrackChange(sid),
    handleSubtitleTrackListChange: (trackList) =>
      notificationRuntime.handleSubtitleTrackListChange(trackList),
    maybeSignalPluginAutoplayReady: (payload, options) =>
      autoplayGate.maybeSignalPluginAutoplayReady(payload, options),
    invalidatePendingAutoplayReadyFallbacks: () =>
      autoplayGate.invalidatePendingAutoplayReadyFallbacks(),
    getAutoPlayReadySignalMediaPath: () => autoplayGate.getAutoPlayReadySignalMediaPath(),
    clearYoutubePlayQuitOnDisconnectArmTimer: () =>
      playbackRuntime.clearYoutubePlayQuitOnDisconnectArmTimer(),
    getQuitOnDisconnectArmed: () => playbackRuntime.getQuitOnDisconnectArmed(),
    isAppOwnedFlowInFlight: () => notificationRuntime.isAppOwnedFlowInFlight(),
  };
}

export { clearYoutubePrimarySubtitleNotificationTimer };
