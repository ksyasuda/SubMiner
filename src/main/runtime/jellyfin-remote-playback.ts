import type { ActiveJellyfinRemotePlaybackState } from './jellyfin-remote-commands';

type MpvClientLike = {
  currentTimePos?: number;
  currentVideoPath?: string;
  requestProperty?: (name: string) => Promise<unknown>;
};

export function secondsToJellyfinTicks(seconds: number, ticksPerSecond: number): number {
  if (!Number.isFinite(seconds)) return 0;
  return Math.max(0, Math.floor(seconds * ticksPerSecond));
}

export function markJellyfinRemotePlaybackLoaded(
  playback: ActiveJellyfinRemotePlaybackState | null,
  path: string,
): void {
  const normalizedPath = path.trim();
  if (playback && normalizedPath) {
    playback.loadedMediaPath = normalizedPath;
  }
}

export function shouldAutoLoadSecondarySubTrackForJellyfinPlayback(
  playback: ActiveJellyfinRemotePlaybackState | null,
  path: string,
): boolean {
  const normalizedPath = path.trim();
  if (!normalizedPath || !playback) {
    return true;
  }
  const loadedMediaPath = playback.loadedMediaPath?.trim() ?? '';
  if (!loadedMediaPath) {
    return false;
  }
  return loadedMediaPath !== normalizedPath;
}

function isMpvPauseEnabled(value: unknown): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (!normalized || normalized === 'no' || normalized === 'false' || normalized === '0') {
      return false;
    }
    return true;
  }
  return false;
}

function normalizeMpvPositionSeconds(value: unknown): number {
  const seconds = Number(value);
  if (!Number.isFinite(seconds)) return 0;
  return Math.max(0, seconds);
}

function getCachedMpvPositionSeconds(client: MpvClientLike | null): number | null {
  if (!client) return null;
  const seconds = Number(client.currentTimePos);
  return Number.isFinite(seconds) ? Math.max(0, seconds) : null;
}

async function readMpvPositionSeconds(client: MpvClientLike | null): Promise<number> {
  const cached = getCachedMpvPositionSeconds(client);
  if (cached !== null) return cached;
  const position = await client?.requestProperty?.('time-pos');
  return normalizeMpvPositionSeconds(position);
}

async function readMpvPositionSecondsOrFallback(
  client: MpvClientLike | null,
  fallback = 0,
): Promise<number> {
  try {
    return await readMpvPositionSeconds(client);
  } catch {
    return fallback;
  }
}

function cacheLastKnownPosition(
  playback: ActiveJellyfinRemotePlaybackState,
  positionSeconds: number,
): void {
  if (!Number.isFinite(positionSeconds)) return;
  if (positionSeconds > 0 || playback.lastKnownPositionSeconds === undefined) {
    playback.lastKnownPositionSeconds = Math.max(0, positionSeconds);
  }
}

function resolveReportablePositionSeconds(
  playback: ActiveJellyfinRemotePlaybackState,
  positionSeconds: number,
): number {
  const normalizedPosition = normalizeMpvPositionSeconds(positionSeconds);
  if (normalizedPosition > 0) return normalizedPosition;
  const cachedPosition = playback.lastKnownPositionSeconds;
  if (typeof cachedPosition === 'number' && Number.isFinite(cachedPosition) && cachedPosition > 0) {
    return cachedPosition;
  }
  return normalizedPosition;
}

function isSeekLikePositionJump(
  previousPositionSeconds: number | null,
  nextPositionSeconds: number,
  thresholdSeconds: number,
): boolean {
  if (previousPositionSeconds === null) return false;
  return Math.abs(nextPositionSeconds - previousPositionSeconds) >= thresholdSeconds;
}

// Jellyfin re-creates a session's NowPlayingItem from any progress report, so a progress
// tick that lands after the stop report leaves the server showing playback forever. The
// tracker lets the stop handler wait for reports that are already in flight. Each report is
// tagged with the playback it describes, so stopping one playback can skip another's reports.
export type JellyfinRemoteReportTracker = {
  track: (report: Promise<void>, playback: ActiveJellyfinRemotePlaybackState | null) => void;
  /** Waits for in-flight reports: every one, or only those tagged with `playback`. */
  settled: (playback?: ActiveJellyfinRemotePlaybackState) => Promise<void>;
};

export function createJellyfinRemoteReportTracker(): JellyfinRemoteReportTracker {
  const active = new Map<Promise<void>, ActiveJellyfinRemotePlaybackState | null>();
  const pending = (playback?: ActiveJellyfinRemotePlaybackState) =>
    [...active].filter(([, owner]) => !playback || owner === playback).map(([report]) => report);
  return {
    track: (report, playback) => {
      active.set(report, playback);
      void report.finally(() => active.delete(report));
    },
    settled: async (playback) => {
      for (let reports = pending(playback); reports.length > 0; reports = pending(playback)) {
        await Promise.allSettled(reports);
      }
    },
  };
}

export type JellyfinRemoteProgressReporterDeps = {
  getActivePlayback: () => ActiveJellyfinRemotePlaybackState | null;
  clearActivePlayback: () => void;
  getMpvClient: () => MpvClientLike | null;
  getNow: () => number;
  getLastProgressAtMs: () => number;
  setLastProgressAtMs: (value: number) => void;
  progressIntervalMs: number;
  ticksPerSecond: number;
  logDebug: (message: string, error: unknown) => void;
  reportTracker?: JellyfinRemoteReportTracker;
};

export function createReportJellyfinRemoteProgressHandler(
  deps: JellyfinRemoteProgressReporterDeps,
) {
  let lastReportedPositionSeconds: number | null = null;

  const report = async (
    playback: ActiveJellyfinRemotePlaybackState | null,
    force: boolean,
  ): Promise<void> => {
    if (!playback) return;
    const now = deps.getNow();
    try {
      const mpvClient = deps.getMpvClient();
      const observedPositionSeconds = await readMpvPositionSeconds(mpvClient);
      cacheLastKnownPosition(playback, observedPositionSeconds);
      const positionSeconds = resolveReportablePositionSeconds(playback, observedPositionSeconds);
      const forceForSeekJump = isSeekLikePositionJump(
        lastReportedPositionSeconds,
        positionSeconds,
        Math.max(2, deps.progressIntervalMs / 1000),
      );
      if (
        !force &&
        !forceForSeekJump &&
        now - deps.getLastProgressAtMs() < deps.progressIntervalMs
      ) {
        return;
      }
      const paused = await mpvClient?.requestProperty?.('pause');
      await playback.reporter.reportProgress({
        itemId: playback.itemId,
        mediaSourceId: playback.mediaSourceId,
        positionTicks: secondsToJellyfinTicks(positionSeconds, deps.ticksPerSecond),
        isPaused: isMpvPauseEnabled(paused),
        playMethod: playback.playMethod,
        audioStreamIndex: playback.audioStreamIndex,
        subtitleStreamIndex: playback.subtitleStreamIndex,
        eventName: 'TimeUpdate',
      });
      lastReportedPositionSeconds = positionSeconds;
      deps.setLastProgressAtMs(now);
    } catch (error) {
      deps.logDebug('Failed to report Jellyfin remote progress', error);
    }
  };

  return async (force = false): Promise<void> => {
    const playback = deps.getActivePlayback();
    const pending = report(playback, force);
    deps.reportTracker?.track(pending, playback);
    await pending;
  };
}

export type JellyfinRemoteStoppedReporterDeps = {
  getActivePlayback: () => ActiveJellyfinRemotePlaybackState | null;
  clearActivePlayback: () => void;
  getMpvClient: () => MpvClientLike | null;
  getNow?: () => number;
  ticksPerSecond: number;
  logDebug: (message: string, error: unknown) => void;
  logWarn?: (message: string) => void;
  reportTracker?: JellyfinRemoteReportTracker;
};

type JellyfinStopReportDeps = {
  ticksPerSecond: number;
  logDebug: (message: string, error: unknown) => void;
  logWarn?: (message: string) => void;
};

async function sendJellyfinStopReports(
  deps: JellyfinStopReportDeps,
  playback: ActiveJellyfinRemotePlaybackState,
  positionSeconds: number,
): Promise<void> {
  const positionTicks = secondsToJellyfinTicks(positionSeconds, deps.ticksPerSecond);
  try {
    await playback.reporter.reportProgress({
      itemId: playback.itemId,
      mediaSourceId: playback.mediaSourceId,
      positionTicks,
      isPaused: false,
      playMethod: playback.playMethod,
      audioStreamIndex: playback.audioStreamIndex,
      subtitleStreamIndex: playback.subtitleStreamIndex,
      eventName: 'TimeUpdate',
    });
  } catch (error) {
    deps.logDebug('Failed to report Jellyfin remote final progress', error);
  }
  const reported = await playback.reporter.reportStopped({
    itemId: playback.itemId,
    mediaSourceId: playback.mediaSourceId,
    positionTicks,
    failed: false,
    playMethod: playback.playMethod,
    audioStreamIndex: playback.audioStreamIndex,
    subtitleStreamIndex: playback.subtitleStreamIndex,
    eventName: 'stop',
  });
  if (reported === false) {
    deps.logWarn?.(
      `Jellyfin did not accept the playback stop report for item ${playback.itemId}; the server may keep showing it as playing.`,
    );
  }
}

export function createReportJellyfinRemoteStoppedHandler(deps: JellyfinRemoteStoppedReporterDeps) {
  return async (): Promise<void> => {
    const playback = deps.getActivePlayback();
    if (!playback) return;
    const withinStartupGrace =
      typeof playback.stopReportsAfterMs === 'number' &&
      Number.isFinite(playback.stopReportsAfterMs) &&
      (deps.getNow?.() ?? Date.now()) < playback.stopReportsAfterMs;
    if (playback.loadedMediaPath === null) {
      // Playback is armed before `loadfile`, so an empty path this soon belongs to the file
      // being replaced and the new item's path is still coming. Later, the load never happened.
      if (!withinStartupGrace) deps.clearActivePlayback();
      return;
    }
    if (withinStartupGrace) return;
    // Clear before any network call so progress ticks fired during the stop find nothing to
    // report, then let reports already in flight finish so none can arrive after the stop.
    deps.clearActivePlayback();
    await deps.reportTracker?.settled();
    try {
      // mpv may have moved to another file while reports drained, and its position would then
      // overwrite this item's resume point. Off its file, the last position observed for it
      // (cached on every time-pos update) stands in.
      const mpvClient = deps.getMpvClient();
      const stillOnItem =
        Boolean(playback.loadedMediaPath) &&
        mpvClient?.currentVideoPath?.trim() === playback.loadedMediaPath?.trim();
      const observedPositionSeconds = stillOnItem
        ? await readMpvPositionSecondsOrFallback(mpvClient)
        : 0;
      const positionSeconds = resolveReportablePositionSeconds(playback, observedPositionSeconds);
      await sendJellyfinStopReports(deps, playback, positionSeconds);
    } catch (error) {
      deps.logDebug('Failed to report Jellyfin remote stop', error);
    }
  };
}

/**
 * Reports the stop of a playback mpv has already left for another playlist entry. mpv's position
 * now belongs to the new file, so the stop carries the last position observed for this one.
 * The caller has already replaced the active playback.
 */
export function createReportSupersededJellyfinPlaybackHandler(
  deps: JellyfinStopReportDeps & { reportTracker?: JellyfinRemoteReportTracker },
) {
  return async (playback: ActiveJellyfinRemotePlaybackState): Promise<void> => {
    // Only this playback's reports: the new one keeps reporting and must not hold the stop.
    await deps.reportTracker?.settled(playback);
    try {
      await sendJellyfinStopReports(deps, playback, playback.lastKnownPositionSeconds ?? 0);
    } catch (error) {
      deps.logDebug('Failed to report Jellyfin remote stop', error);
    }
  };
}
