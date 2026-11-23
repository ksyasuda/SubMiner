import type { JellyfinAuthSession, JellyfinPlaybackPlan } from '../../core/services/jellyfin';
import type { JellyfinConfig } from '../../types';
import type { MpvRuntimeClientLike } from '../../core/services/mpv';
import {
  createJellyfinEpisodePlaylist,
  type JellyfinEpisodePlaylistSettings,
} from './jellyfin-episode-playlist';
import type {
  ActiveJellyfinRemotePlaybackState,
  JellyfinPlaybackReporterLike,
} from './jellyfin-remote-commands';

type JellyfinClientInfo = {
  clientName: string;
  clientVersion: string;
  deviceId: string;
};

type JellyfinPlaybackContext = {
  session: JellyfinAuthSession;
  clientInfo: JellyfinClientInfo;
  jellyfinConfig: JellyfinConfig;
};

type ActivePlaybackState = {
  itemId: string;
  reporter: JellyfinPlaybackReporterLike;
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

/** A resolved Jellyfin item in the form mpv loads it. */
type JellyfinPreparedPlayback = {
  itemId: string;
  seriesId: string | null;
  title: string;
  plan: JellyfinPlaybackPlan;
  playbackUrl: string;
  loadfileOptions: string;
  startTimeTicks: number;
  startSeconds: number;
  playMethod: 'DirectPlay' | 'Transcode';
};

// mpv's `%<bytes>%value` quoting takes any value verbatim, commas and colons included.
function quoteMpvOptionValue(value: string): string {
  return `%${Buffer.byteLength(value, 'utf8')}%${value}`;
}

function buildJellyfinLoadfileOptions(
  plan: JellyfinPlaybackPlan,
  startSeconds: number,
  mediaTitle?: string,
): string {
  const options = [...JELLYFIN_LOADFILE_SUBTITLE_SUPPRESSION_OPTIONS];
  if (plan.mode === 'direct' && startSeconds > 0) {
    options.push(`start=${startSeconds}`);
  }
  if (mediaTitle) {
    options.push(`force-media-title=${quoteMpvOptionValue(mediaTitle)}`);
  }
  return options.join(',');
}

function prepareJellyfinPlayback(
  itemId: string,
  plan: JellyfinPlaybackPlan,
  convertTicksToSeconds: (ticks: number) => number,
  options: {
    startTimeTicksOverride?: number;
    fallbackToPlanStartTimeOnZeroOverride?: boolean;
    // Playlist entries carry their title per file; an explicit play sets it before loadfile.
    perFileMediaTitle?: boolean;
  } = {},
): JellyfinPreparedPlayback {
  const startTimeTicks = resolveEffectiveStartTimeTicks(
    plan.startTimeTicks,
    options.startTimeTicksOverride,
    options.fallbackToPlanStartTimeOnZeroOverride,
  );
  const startSeconds = startTimeTicks > 0 ? Math.max(0, convertTicksToSeconds(startTimeTicks)) : 0;
  const playbackUrlBase =
    plan.mode === 'direct'
      ? stripStartTimeTicksFromPlaybackUrl(plan.url)
      : applyStartTimeTicksToPlaybackUrl(plan.url, startTimeTicks);
  return {
    itemId,
    seriesId: plan.seriesId,
    title: plan.title,
    plan,
    playbackUrl: stripManagedSubtitleStreamFromPlaybackUrl(playbackUrlBase),
    loadfileOptions: buildJellyfinLoadfileOptions(
      plan,
      startSeconds,
      options.perFileMediaTitle ? plan.title : undefined,
    ),
    startTimeTicks,
    startSeconds,
    playMethod: plan.mode === 'direct' ? 'DirectPlay' : 'Transcode',
  };
}

export type PlayJellyfinItemInMpvHandler = ((params: {
  session: JellyfinAuthSession;
  clientInfo: JellyfinClientInfo;
  jellyfinConfig: JellyfinConfig;
  itemId: string;
  audioStreamIndex?: number | null;
  subtitleStreamIndex?: number | null;
  startTimeTicksOverride?: number;
  fallbackToPlanStartTimeOnZeroOverride?: boolean;
  setQuitOnDisconnectArm?: boolean;
}) => Promise<void>) & {
  /**
   * Takes over a queued neighbor episode that mpv started from its playlist, given mpv's new
   * `path`. Returns false when the path is not one, or is already the active playback. Runs
   * synchronously up to recording stats metadata, so call it before the immersion tracker sees
   * the path.
   */
  adoptQueuedEpisode: (path: string) => boolean;
  /**
   * Called on mpv time-pos updates. The first one for a just-started item queues its neighbor
   * episodes, so that work never competes with opening the stream.
   */
  handlePlaybackTimePos: () => void;
};

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
  getActivePlayback: () => ActiveJellyfinRemotePlaybackState | null;
  setActivePlayback: (state: ActivePlaybackState) => void;
  setLastProgressAtMs: (value: number) => void;
  // Reports go out as the device and account that started this playback, for CLI play and
  // casting alike, so Jellyfin keeps the resume point and played state either way.
  createPlaybackReporter: (params: {
    session: JellyfinAuthSession;
    clientInfo: JellyfinClientInfo;
  }) => JellyfinPlaybackReporterLike;
  /** Reports the stop of a playback mpv left for a queued neighbor episode. */
  reportPlaybackSuperseded: (playback: ActiveJellyfinRemotePlaybackState) => Promise<void>;
  /** Resolves once no Jellyfin episode start or stop report is in flight, whoever sent it. */
  waitForTransitionReports: () => Promise<void>;
  /** Tracks a start report, so the playback's own stop and later starts wait for it. */
  trackPlaybackStart: (report: Promise<void>, playback: ActiveJellyfinRemotePlaybackState) => void;
  /** Read on every queue run, so edits apply from the next episode change. */
  getEpisodePlaylistSettings: () => JellyfinEpisodePlaylistSettings;
  listSeriesEpisodeIds: (params: {
    session: JellyfinAuthSession;
    clientInfo: JellyfinClientInfo;
    seriesId: string;
  }) => Promise<string[]>;
  readMpvPlaylist: () => Promise<unknown>;
  showMpvOsd: (text: string) => void;
  recordJellyfinPlaybackMetadata?: (
    metadata: JellyfinPlaybackStatsMetadata,
  ) => void | Promise<void>;
  updateCurrentMediaTitle?: (title: string) => void | Promise<void>;
  logDebug: (message: string, error?: unknown) => void;
}): PlayJellyfinItemInMpvHandler {
  const episodePlaylist = createJellyfinEpisodePlaylist<
    JellyfinPlaybackContext,
    JellyfinPreparedPlayback
  >({
    getSettings: () => deps.getEpisodePlaylistSettings(),
    listSeriesEpisodeIds: (context, seriesId) =>
      deps.listSeriesEpisodeIds({
        session: context.session,
        clientInfo: context.clientInfo,
        seriesId,
      }),
    prepareEntry: async (context, itemId) =>
      prepareJellyfinPlayback(
        itemId,
        await deps.resolvePlaybackPlan({ ...context, itemId }),
        deps.convertTicksToSeconds,
        { perFileMediaTitle: true },
      ),
    readMpvPlaylist: () => deps.readMpvPlaylist(),
    sendMpvCommand: (command) => deps.sendMpvCommand(command),
    logDebug: (message, error) => deps.logDebug(message, error),
  });

  // Records stats metadata and makes `prepared` the active playback. Synchronous, so a caller
  // reacting to mpv's path change finishes it before the immersion tracker sees that path.
  const beginPlayback = (
    context: JellyfinPlaybackContext,
    prepared: JellyfinPreparedPlayback,
    loadedMediaPath: string | null,
  ): JellyfinPlaybackReporterLike => {
    const { plan } = prepared;
    runBestEffortPlaybackHook(() => deps.updateCurrentMediaTitle?.(plan.title));
    runBestEffortPlaybackHook(() =>
      deps.recordJellyfinPlaybackMetadata?.({
        mediaPath: prepared.playbackUrl,
        displayTitle: plan.title,
        itemTitle: plan.itemTitle,
        seriesTitle: plan.seriesTitle,
        seasonNumber: plan.seasonNumber,
        episodeNumber: plan.episodeNumber,
        itemId: prepared.itemId,
      }),
    );
    const reporter = deps.createPlaybackReporter({
      session: context.session,
      clientInfo: context.clientInfo,
    });
    deps.setActivePlayback({
      itemId: prepared.itemId,
      reporter,
      mediaSourceId: undefined,
      audioStreamIndex: plan.audioStreamIndex,
      subtitleStreamIndex: plan.subtitleStreamIndex,
      playMethod: prepared.playMethod,
      loadedMediaPath,
      lastKnownPositionSeconds: prepared.startSeconds > 0 ? prepared.startSeconds : undefined,
    });
    deps.setLastProgressAtMs(0);
    return reporter;
  };

  // The item whose neighbors get queued once mpv reports it playing.
  let pendingNeighborQueue: {
    context: JellyfinPlaybackContext;
    prepared: JellyfinPreparedPlayback;
  } | null = null;

  // Runs once mpv has the item: subtitles, overlay, and the Jellyfin start report. Its
  // neighbors are queued later, from the first time-pos update.
  const finishPlaybackStartup = async (
    context: JellyfinPlaybackContext,
    prepared: JellyfinPreparedPlayback,
    reporter: JellyfinPlaybackReporterLike,
  ): Promise<void> => {
    const { plan } = prepared;
    // Each start makes a new reporter, so it identifies this playback even when the same
    // episode comes back. A later switch that took over does its own reporting.
    const ownsPlayback = () => deps.getActivePlayback()?.reporter === reporter;
    await awaitBestEffortPlaybackHook(() =>
      deps.preloadExternalSubtitles({
        session: context.session,
        clientInfo: context.clientInfo,
        itemId: prepared.itemId,
      }),
    );
    if (!ownsPlayback()) return;
    deps.showVisibleOverlay();
    pendingNeighborQueue = { context, prepared };

    // Jellyfin clears a session's now-playing item on any stop and keeps whichever start lands
    // last, so this start waits for every start and stop in flight, including the generic stop
    // mpv's empty path between files can trigger.
    await deps.waitForTransitionReports();
    const playback = deps.getActivePlayback();
    if (playback?.reporter !== reporter) return;
    const started = reporter.reportPlaying({
      itemId: prepared.itemId,
      mediaSourceId: undefined,
      playMethod: prepared.playMethod,
      positionTicks: prepared.startTimeTicks,
      isPaused: false,
      audioStreamIndex: plan.audioStreamIndex,
      subtitleStreamIndex: plan.subtitleStreamIndex,
      eventName: 'start',
    });
    deps.trackPlaybackStart(
      started.then(
        () => undefined,
        (error: unknown) => deps.logDebug('Failed to report Jellyfin playback start', error),
      ),
      playback,
    );
    deps.showMpvOsd(`Jellyfin ${plan.mode}: ${plan.title}`);
  };

  const play = async (params: Parameters<PlayJellyfinItemInMpvHandler>[0]): Promise<void> => {
    const connected = await deps.ensureMpvConnectedForPlayback();
    const mpvClient = deps.getMpvClient();
    if (!connected || !mpvClient) {
      throw new Error(
        'MPV not connected and auto-launch failed. Check mpv.executablePath or ensure mpv is available in PATH.',
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
    const prepared = prepareJellyfinPlayback(params.itemId, plan, deps.convertTicksToSeconds, {
      startTimeTicksOverride: params.startTimeTicksOverride,
      fallbackToPlanStartTimeOnZeroOverride: params.fallbackToPlanStartTimeOnZeroOverride,
    });
    const context: JellyfinPlaybackContext = {
      session: params.session,
      clientInfo: params.clientInfo,
      jellyfinConfig: params.jellyfinConfig,
    };
    // `loadfile replace` empties mpv's playlist, taking the old neighbors with it.
    episodePlaylist.reset();
    pendingNeighborQueue = null;
    const reporter = beginPlayback(context, prepared, null);
    deps.sendMpvCommand(['script-message', 'subminer-managed-subtitles-loading']);
    // Set mpv's title before loadfile can emit a URL-derived media-title event.
    deps.sendMpvCommand(['set_property', 'force-media-title', plan.title]);
    deps.sendMpvCommand([
      'loadfile',
      prepared.playbackUrl,
      'replace',
      -1,
      prepared.loadfileOptions,
    ]);
    if (params.setQuitOnDisconnectArm !== false) {
      deps.armQuitOnDisconnect();
    }

    await finishPlaybackStartup(context, prepared, reporter);
  };

  const adoptQueuedEpisode = (path: string): boolean => {
    const queued = episodePlaylist.find(path);
    if (!queued) return false;
    const previous = deps.getActivePlayback();
    if (previous?.itemId === queued.entry.itemId) return false;

    const reporter = beginPlayback(queued.context, queued.entry, path);
    if (previous) {
      void deps.reportPlaybackSuperseded(previous).catch((error: unknown) => {
        deps.logDebug('Failed to report the previous Jellyfin episode stopped', error);
      });
    }
    void finishPlaybackStartup(queued.context, queued.entry, reporter).catch((error: unknown) => {
      deps.logDebug('Failed to start the queued Jellyfin episode', error);
    });
    return true;
  };

  const handlePlaybackTimePos = (): void => {
    const pending = pendingNeighborQueue;
    if (!pending) return;
    const active = deps.getActivePlayback();
    // time-pos can still belong to the file mpv is leaving until it reports the new path.
    if (
      active?.itemId !== pending.prepared.itemId ||
      active.loadedMediaPath !== pending.prepared.playbackUrl
    ) {
      return;
    }
    pendingNeighborQueue = null;
    void episodePlaylist.queueNeighbors(pending.context, pending.prepared);
  };

  return Object.assign(play, { adoptQueuedEpisode, handlePlaybackTimePos });
}
