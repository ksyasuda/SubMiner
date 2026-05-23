import type { JellyfinAuthSession, JellyfinPlaybackPlan } from '../../core/services/jellyfin';
import type { JellyfinConfig } from '../../types';
import type { MpvRuntimeClientLike } from '../../core/services/mpv';

type JellyfinClientInfo = {
  clientName: string;
  clientVersion: string;
  deviceId: string;
};

type ActivePlaybackState = {
  itemId: string;
  mediaSourceId: undefined;
  audioStreamIndex?: number | null;
  subtitleStreamIndex?: number | null;
  playMethod: 'DirectPlay' | 'Transcode';
  loadedMediaPath?: string | null;
  stopReportsAfterMs?: number;
};

export type JellyfinPlaybackStatsMetadata = {
  mediaPath: string;
  displayTitle: string;
  itemTitle: string;
  seriesTitle: string | null;
  seasonNumber: number | null;
  episodeNumber: number | null;
  itemId: string;
};

function applyStartTimeTicksToPlaybackUrl(url: string, startTimeTicksOverride?: number): string {
  if (typeof startTimeTicksOverride !== 'number') return url;
  try {
    const resolved = new URL(url);
    if (startTimeTicksOverride > 0) {
      resolved.searchParams.set('StartTimeTicks', String(Math.max(0, startTimeTicksOverride)));
    } else {
      resolved.searchParams.delete('StartTimeTicks');
    }
    return resolved.toString();
  } catch {
    return url;
  }
}

export function createPlayJellyfinItemInMpvHandler(deps: {
  ensureMpvConnectedForPlayback: () => Promise<boolean>;
  getMpvClient: () => MpvRuntimeClientLike | null;
  resolvePlaybackPlan: (params: {
    session: JellyfinAuthSession;
    clientInfo: JellyfinClientInfo;
    jellyfinConfig: JellyfinConfig;
    itemId: string;
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
  }) => Promise<JellyfinPlaybackPlan>;
  applyJellyfinMpvDefaults: (mpvClient: MpvRuntimeClientLike) => void;
  showVisibleOverlay: () => void;
  sendMpvCommand: (command: Array<string | number>) => void;
  armQuitOnDisconnect: () => void;
  schedule: (callback: () => void, delayMs: number) => void;
  convertTicksToSeconds: (ticks: number) => number;
  preloadExternalSubtitles: (params: {
    session: JellyfinAuthSession;
    clientInfo: JellyfinClientInfo;
    itemId: string;
  }) => void;
  setActivePlayback: (state: ActivePlaybackState) => void;
  setLastProgressAtMs: (value: number) => void;
  reportPlaying: (payload: {
    itemId: string;
    mediaSourceId: undefined;
    playMethod: 'DirectPlay' | 'Transcode';
    positionTicks?: number;
    isPaused?: boolean;
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
    eventName: 'start';
  }) => void;
  showMpvOsd: (text: string) => void;
  recordJellyfinPlaybackMetadata?: (metadata: JellyfinPlaybackStatsMetadata) => void;
  updateCurrentMediaTitle?: (title: string) => void;
}) {
  return async (params: {
    session: JellyfinAuthSession;
    clientInfo: JellyfinClientInfo;
    jellyfinConfig: JellyfinConfig;
    itemId: string;
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
    startTimeTicksOverride?: number;
    setQuitOnDisconnectArm?: boolean;
  }): Promise<void> => {
    const connected = await deps.ensureMpvConnectedForPlayback();
    const mpvClient = deps.getMpvClient();
    if (!connected || !mpvClient) {
      throw new Error(
        'MPV not connected and auto-launch failed. Ensure mpv is installed and available in PATH.',
      );
    }

    const plan = await deps.resolvePlaybackPlan({
      session: params.session,
      clientInfo: params.clientInfo,
      jellyfinConfig: params.jellyfinConfig,
      itemId: params.itemId,
      audioStreamIndex: params.audioStreamIndex,
      subtitleStreamIndex: params.subtitleStreamIndex,
    });

    deps.applyJellyfinMpvDefaults(mpvClient);
    deps.sendMpvCommand(['set_property', 'sub-auto', 'no']);
    const playbackUrl = applyStartTimeTicksToPlaybackUrl(plan.url, params.startTimeTicksOverride);
    const playMethod = plan.mode === 'direct' ? 'DirectPlay' : 'Transcode';
    try {
      deps.updateCurrentMediaTitle?.(plan.title);
      deps.recordJellyfinPlaybackMetadata?.({
        mediaPath: playbackUrl,
        displayTitle: plan.title,
        itemTitle: plan.itemTitle,
        seriesTitle: plan.seriesTitle,
        seasonNumber: plan.seasonNumber,
        episodeNumber: plan.episodeNumber,
        itemId: params.itemId,
      });
    } catch {
      // Best-effort metadata/title hooks must not block playback startup.
    }
    deps.setActivePlayback({
      itemId: params.itemId,
      mediaSourceId: undefined,
      audioStreamIndex: plan.audioStreamIndex,
      subtitleStreamIndex: plan.subtitleStreamIndex,
      playMethod,
      loadedMediaPath: null,
    });
    deps.setLastProgressAtMs(0);
    deps.sendMpvCommand(['loadfile', playbackUrl, 'replace']);
    if (params.setQuitOnDisconnectArm !== false) {
      deps.armQuitOnDisconnect();
    }
    deps.sendMpvCommand(['set_property', 'force-media-title', plan.title]);
    deps.sendMpvCommand(['set_property', 'sid', 'no']);

    const startTimeTicks =
      typeof params.startTimeTicksOverride === 'number'
        ? Math.max(0, params.startTimeTicksOverride)
        : plan.startTimeTicks;
    if (startTimeTicks > 0) {
      deps.sendMpvCommand(['seek', deps.convertTicksToSeconds(startTimeTicks), 'absolute+exact']);
    }

    deps.showVisibleOverlay();
    deps.preloadExternalSubtitles({
      session: params.session,
      clientInfo: params.clientInfo,
      itemId: params.itemId,
    });

    deps.reportPlaying({
      itemId: params.itemId,
      mediaSourceId: undefined,
      playMethod,
      positionTicks: startTimeTicks,
      isPaused: false,
      audioStreamIndex: plan.audioStreamIndex,
      subtitleStreamIndex: plan.subtitleStreamIndex,
      eventName: 'start',
    });
    deps.showMpvOsd(`Jellyfin ${plan.mode}: ${plan.title}`);
  };
}
