import type { ActiveJellyfinRemotePlaybackState } from './jellyfin-remote-commands';

type JellyfinRemoteSessionLike = {
  isConnected: () => boolean;
  reportProgress: (payload: {
    itemId: string;
    mediaSourceId?: string;
    positionTicks: number;
    isPaused: boolean;
    playMethod: 'DirectPlay' | 'Transcode';
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
    eventName: 'TimeUpdate';
  }) => Promise<unknown>;
  reportStopped: (payload: {
    itemId: string;
    mediaSourceId?: string;
    positionTicks?: number;
    failed?: boolean;
    playMethod: 'DirectPlay' | 'Transcode';
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
    eventName: 'stop';
  }) => Promise<unknown>;
};

type MpvClientLike = {
  currentTimePos?: number;
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

export type JellyfinRemoteProgressReporterDeps = {
  getActivePlayback: () => ActiveJellyfinRemotePlaybackState | null;
  clearActivePlayback: () => void;
  getSession: () => JellyfinRemoteSessionLike | null;
  getMpvClient: () => MpvClientLike | null;
  getNow: () => number;
  getLastProgressAtMs: () => number;
  setLastProgressAtMs: (value: number) => void;
  progressIntervalMs: number;
  ticksPerSecond: number;
  logDebug: (message: string, error: unknown) => void;
};

export function createReportJellyfinRemoteProgressHandler(
  deps: JellyfinRemoteProgressReporterDeps,
) {
  let lastReportedPositionSeconds: number | null = null;

  return async (force = false): Promise<void> => {
    const playback = deps.getActivePlayback();
    if (!playback) return;
    const session = deps.getSession();
    // Timeline posts are HTTP requests; keep them flowing while the remote websocket reconnects.
    if (!session) return;
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
      await session.reportProgress({
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
}

export type JellyfinRemoteStoppedReporterDeps = {
  getActivePlayback: () => ActiveJellyfinRemotePlaybackState | null;
  clearActivePlayback: () => void;
  getSession: () => JellyfinRemoteSessionLike | null;
  getMpvClient: () => MpvClientLike | null;
  getNow?: () => number;
  ticksPerSecond: number;
  logDebug: (message: string, error: unknown) => void;
};

export function createReportJellyfinRemoteStoppedHandler(deps: JellyfinRemoteStoppedReporterDeps) {
  return async (): Promise<void> => {
    const playback = deps.getActivePlayback();
    if (!playback) return;
    if (playback.loadedMediaPath === null) {
      deps.clearActivePlayback();
      return;
    }
    if (
      typeof playback.stopReportsAfterMs === 'number' &&
      Number.isFinite(playback.stopReportsAfterMs) &&
      (deps.getNow?.() ?? Date.now()) < playback.stopReportsAfterMs
    ) {
      return;
    }
    const session = deps.getSession();
    // Timeline posts are HTTP requests; keep them flowing while the remote websocket reconnects.
    if (!session) {
      deps.clearActivePlayback();
      return;
    }
    try {
      const observedPositionSeconds = await readMpvPositionSecondsOrFallback(deps.getMpvClient());
      const positionSeconds = resolveReportablePositionSeconds(playback, observedPositionSeconds);
      const positionTicks = secondsToJellyfinTicks(positionSeconds, deps.ticksPerSecond);
      try {
        await session.reportProgress({
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
      await session.reportStopped({
        itemId: playback.itemId,
        mediaSourceId: playback.mediaSourceId,
        positionTicks,
        failed: false,
        playMethod: playback.playMethod,
        audioStreamIndex: playback.audioStreamIndex,
        subtitleStreamIndex: playback.subtitleStreamIndex,
        eventName: 'stop',
      });
    } catch (error) {
      deps.logDebug('Failed to report Jellyfin remote stop', error);
    } finally {
      deps.clearActivePlayback();
    }
  };
}
