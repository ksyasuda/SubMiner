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
  lastKnownPositionSeconds?: number;
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

const JELLYFIN_LOADFILE_SUBTITLE_SUPPRESSION_OPTIONS = [
  'sid=no',
  'secondary-sid=no',
  'sub-auto=no',
  'sub-visibility=no',
  'secondary-sub-visibility=no',
];

function runBestEffortPlaybackHook(callback: () => void | Promise<void>): void {
  try {
    void Promise.resolve(callback()).catch(() => {});
  } catch {
    // Best-effort metadata/title hooks must not block playback startup.
  }
}

async function awaitBestEffortPlaybackHook(callback: () => void | Promise<void>): Promise<void> {
  try {
    await Promise.resolve(callback());
  } catch {
    // Best-effort startup hooks must not block playback startup.
  }
}

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

function stripStartTimeTicksFromPlaybackUrl(url: string): string {
  try {
    const resolved = new URL(url);
    resolved.searchParams.delete('StartTimeTicks');
    return resolved.toString();
  } catch {
    return url;
  }
}

function stripManagedSubtitleStreamFromPlaybackUrl(url: string): string {
  try {
    const resolved = new URL(url);
    resolved.searchParams.delete('SubtitleStreamIndex');
    return resolved.toString();
  } catch {
    return url;
  }
}

function resolveEffectiveStartTimeTicks(
  planStartTimeTicks: number,
  startTimeTicksOverride?: number,
  fallbackToPlanStartTimeOnZeroOverride = false,
) {
  if (typeof startTimeTicksOverride === 'number' && startTimeTicksOverride > 0) {
    return Math.max(0, startTimeTicksOverride);
  }
  if (typeof startTimeTicksOverride === 'number') {
    return fallbackToPlanStartTimeOnZeroOverride ? Math.max(0, planStartTimeTicks) : 0;
  }
  return Math.max(0, planStartTimeTicks);
}

function buildJellyfinLoadfileOptions(plan: JellyfinPlaybackPlan, startSeconds: number): string {
  const options = [...JELLYFIN_LOADFILE_SUBTITLE_SUPPRESSION_OPTIONS];
  if (plan.mode === 'direct' && startSeconds > 0) {
    options.push(`start=${startSeconds}`);
  }
  return options.join(',');
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
  }) => void | Promise<void>;
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
  recordJellyfinPlaybackMetadata?: (
    metadata: JellyfinPlaybackStatsMetadata,
  ) => void | Promise<void>;
  updateCurrentMediaTitle?: (title: string) => void | Promise<void>;
}) {
  return async (params: {
    session: JellyfinAuthSession;
    clientInfo: JellyfinClientInfo;
    jellyfinConfig: JellyfinConfig;
    itemId: string;
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
    startTimeTicksOverride?: number;
    fallbackToPlanStartTimeOnZeroOverride?: boolean;
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
    deps.sendMpvCommand(['set_property', 'sid', 'no']);
    deps.sendMpvCommand(['set_property', 'secondary-sid', 'no']);
    deps.sendMpvCommand(['set_property', 'sub-visibility', 'no']);
    deps.sendMpvCommand(['set_property', 'secondary-sub-visibility', 'no']);
    const startTimeTicks = resolveEffectiveStartTimeTicks(
      plan.startTimeTicks,
      params.startTimeTicksOverride,
      params.fallbackToPlanStartTimeOnZeroOverride,
    );
    const startSeconds =
      startTimeTicks > 0 ? Math.max(0, deps.convertTicksToSeconds(startTimeTicks)) : 0;
    const playbackUrlBase =
      plan.mode === 'direct'
        ? stripStartTimeTicksFromPlaybackUrl(plan.url)
        : applyStartTimeTicksToPlaybackUrl(plan.url, startTimeTicks);
    const playbackUrl = stripManagedSubtitleStreamFromPlaybackUrl(playbackUrlBase);
    const loadfileOptions = buildJellyfinLoadfileOptions(plan, startSeconds);
    const playMethod = plan.mode === 'direct' ? 'DirectPlay' : 'Transcode';
    runBestEffortPlaybackHook(() => deps.updateCurrentMediaTitle?.(plan.title));
    runBestEffortPlaybackHook(() =>
      deps.recordJellyfinPlaybackMetadata?.({
        mediaPath: playbackUrl,
        displayTitle: plan.title,
        itemTitle: plan.itemTitle,
        seriesTitle: plan.seriesTitle,
        seasonNumber: plan.seasonNumber,
        episodeNumber: plan.episodeNumber,
        itemId: params.itemId,
      }),
    );
    deps.setActivePlayback({
      itemId: params.itemId,
      mediaSourceId: undefined,
      audioStreamIndex: plan.audioStreamIndex,
      subtitleStreamIndex: plan.subtitleStreamIndex,
      playMethod,
      loadedMediaPath: null,
      lastKnownPositionSeconds: startSeconds > 0 ? startSeconds : undefined,
    });
    deps.setLastProgressAtMs(0);
    deps.sendMpvCommand(['script-message', 'subminer-managed-subtitles-loading']);
    deps.sendMpvCommand(['loadfile', playbackUrl, 'replace', -1, loadfileOptions]);
    if (params.setQuitOnDisconnectArm !== false) {
      deps.armQuitOnDisconnect();
    }
    deps.sendMpvCommand(['set_property', 'force-media-title', plan.title]);

    await awaitBestEffortPlaybackHook(() =>
      deps.preloadExternalSubtitles({
        session: params.session,
        clientInfo: params.clientInfo,
        itemId: params.itemId,
      }),
    );
    deps.showVisibleOverlay();

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
