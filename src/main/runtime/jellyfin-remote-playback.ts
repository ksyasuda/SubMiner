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
    eventName: 'timeupdate';
  }) => Promise<unknown>;
  reportStopped: (payload: {
    itemId: string;
    mediaSourceId?: string;
    playMethod: 'DirectPlay' | 'Transcode';
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
    eventName: 'stop';
  }) => Promise<unknown>;
};

type MpvClientLike = {
  requestProperty: (name: string) => Promise<unknown>;
};

export function secondsToJellyfinTicks(seconds: number, ticksPerSecond: number): number {
  if (!Number.isFinite(seconds)) return 0;
  return Math.max(0, Math.floor(seconds * ticksPerSecond));
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
  return async (force = false): Promise<void> => {
    const playback = deps.getActivePlayback();
    if (!playback) return;
    const session = deps.getSession();
    if (!session || !session.isConnected()) return;
    const now = deps.getNow();
    if (!force && now - deps.getLastProgressAtMs() < deps.progressIntervalMs) {
      return;
    }
    try {
      const mpvClient = deps.getMpvClient();
      const position = await mpvClient?.requestProperty('time-pos');
      const paused = await mpvClient?.requestProperty('pause');
      await session.reportProgress({
        itemId: playback.itemId,
        mediaSourceId: playback.mediaSourceId,
        positionTicks: secondsToJellyfinTicks(Number(position) || 0, deps.ticksPerSecond),
        isPaused: paused === true,
        playMethod: playback.playMethod,
        audioStreamIndex: playback.audioStreamIndex,
        subtitleStreamIndex: playback.subtitleStreamIndex,
        eventName: 'timeupdate',
      });
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
  logDebug: (message: string, error: unknown) => void;
};

export function createReportJellyfinRemoteStoppedHandler(deps: JellyfinRemoteStoppedReporterDeps) {
  return async (): Promise<void> => {
    const playback = deps.getActivePlayback();
    if (!playback) return;
    const session = deps.getSession();
    if (!session || !session.isConnected()) {
      deps.clearActivePlayback();
      return;
    }
    try {
      await session.reportStopped({
        itemId: playback.itemId,
        mediaSourceId: playback.mediaSourceId,
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
