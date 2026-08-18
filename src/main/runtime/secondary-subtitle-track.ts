import type { SubtitleCue } from '../../types/subtitle';

type SecondarySubtitleMpvClient = {
  connected?: boolean;
  requestProperty: (name: string) => Promise<unknown>;
};

type ResolvedSubtitleSource = {
  path: string;
  sourceKey: string;
  cleanup?: () => Promise<void>;
};

type SecondarySubtitleSourceInput = {
  currentExternalFilenameRaw: unknown;
  currentTrackRaw: unknown;
  trackListRaw: unknown;
  sidRaw: unknown;
  videoPath: string;
  allowSelectedFallback?: boolean;
};

const DEFAULT_REFRESH_DELAY_MS = 500;

function finiteNumber(value: unknown, fallback = 0): number {
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function trackId(value: unknown): number | null {
  if (typeof value !== 'number' && typeof value !== 'string') return null;
  const number = typeof value === 'number' ? value : Number(value.trim());
  return Number.isInteger(number) ? number : null;
}

function buildSelectedTrackIdentity(
  trackListRaw: unknown,
  sidRaw: unknown,
  videoPath: string,
): string | null {
  if (!Array.isArray(trackListRaw)) return null;
  const sid = trackId(sidRaw);
  if (sid === null) return null;

  const selectedTrack = trackListRaw.find((entry: unknown) => {
    if (!entry || typeof entry !== 'object') return false;
    const track = entry as Record<string, unknown>;
    return track.type === 'sub' && trackId(track.id) === sid;
  }) as Record<string, unknown> | undefined;
  if (!selectedTrack) return null;

  return JSON.stringify([
    videoPath,
    sid,
    selectedTrack.external === true,
    selectedTrack['external-filename'] ?? null,
    trackId(selectedTrack['ff-index']),
  ]);
}

export function findActiveSubtitleText(cues: readonly SubtitleCue[], timeSeconds: number): string {
  if (!Number.isFinite(timeSeconds)) return '';

  const seen = new Set<string>();
  const activeText: string[] = [];
  for (const cue of cues) {
    if (cue.startTime > timeSeconds || cue.endTime <= timeSeconds) continue;
    const text = cue.text.trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    activeText.push(text);
  }
  return activeText.join('\n');
}

export function createSecondarySubtitleTrackController(deps: {
  getMpvClient: () => SecondarySubtitleMpvClient | null;
  getCurrentTimePos: () => number;
  resolveSubtitleSource: (
    input: SecondarySubtitleSourceInput,
  ) => Promise<ResolvedSubtitleSource | null>;
  loadSubtitleSourceText: (source: string) => Promise<string>;
  parseSubtitleCues: (content: string, filename: string) => SubtitleCue[];
  setCurrentSecondaryText: (text: string) => void;
  broadcastSecondaryText: (text: string) => void;
  logDebug?: (message: string) => void;
  logWarn?: (message: string, error: unknown) => void;
}) {
  let parsedCues: SubtitleCue[] | null = null;
  let parsedSourceKey: string | null = null;
  let parsedTrackIdentity: string | null = null;
  let secondaryDelaySeconds = 0;
  let lastLiveText = '';
  let lastBroadcastText: string | null = null;
  let refreshGeneration = 0;
  let refreshTimer: ReturnType<typeof setTimeout> | null = null;

  const publish = (text: string): void => {
    deps.setCurrentSecondaryText(text);
    if (text === lastBroadcastText) return;
    lastBroadcastText = text;
    deps.broadcastSecondaryText(text);
  };

  const resolveAtTime = (timeSeconds: number): string => {
    if (!parsedCues) return lastLiveText;
    return findActiveSubtitleText(parsedCues, timeSeconds - secondaryDelaySeconds);
  };

  const useLiveFallback = (): void => {
    parsedCues = null;
    parsedSourceKey = null;
    parsedTrackIdentity = null;
    publish(lastLiveText);
  };

  const refresh = async (): Promise<void> => {
    const generation = ++refreshGeneration;
    const client = deps.getMpvClient();
    if (!client?.connected) {
      useLiveFallback();
      return;
    }

    let resolvedSource: ResolvedSubtitleSource | null = null;
    try {
      const [secondarySid, trackList, videoPathRaw, secondaryDelayRaw] = await Promise.all([
        client.requestProperty('secondary-sid').catch(() => null),
        client.requestProperty('track-list').catch(() => null),
        client.requestProperty('path').catch(() => null),
        client.requestProperty('secondary-sub-delay').catch(() => 0),
      ]);
      if (generation !== refreshGeneration) return;

      const videoPath = typeof videoPathRaw === 'string' ? videoPathRaw.trim() : '';
      if (!videoPath || secondarySid === null || secondarySid === 'no') {
        useLiveFallback();
        return;
      }

      secondaryDelaySeconds = finiteNumber(secondaryDelayRaw);
      const selectedTrackIdentity = buildSelectedTrackIdentity(trackList, secondarySid, videoPath);
      if (selectedTrackIdentity && selectedTrackIdentity === parsedTrackIdentity && parsedCues) {
        publish(resolveAtTime(deps.getCurrentTimePos()));
        return;
      }

      resolvedSource = await deps.resolveSubtitleSource({
        currentExternalFilenameRaw: null,
        currentTrackRaw: null,
        trackListRaw: trackList,
        sidRaw: secondarySid,
        videoPath,
        allowSelectedFallback: false,
      });
      if (generation !== refreshGeneration) return;
      if (!resolvedSource) {
        deps.logDebug?.('[secondary-subtitle-track] selected source is not readable');
        useLiveFallback();
        return;
      }

      if (resolvedSource.sourceKey === parsedSourceKey && parsedCues) {
        parsedTrackIdentity = selectedTrackIdentity;
        publish(resolveAtTime(deps.getCurrentTimePos()));
        return;
      }

      const content = await deps.loadSubtitleSourceText(resolvedSource.path);
      const cues = deps.parseSubtitleCues(content, resolvedSource.path);
      if (generation !== refreshGeneration) return;
      if (cues.length === 0) {
        deps.logDebug?.('[secondary-subtitle-track] selected source contained no parsed cues');
        useLiveFallback();
        return;
      }

      parsedCues = cues;
      parsedSourceKey = resolvedSource.sourceKey;
      parsedTrackIdentity = selectedTrackIdentity;
      publish(resolveAtTime(deps.getCurrentTimePos()));
    } catch (error) {
      if (generation !== refreshGeneration) return;
      deps.logWarn?.('[secondary-subtitle-track] failed to parse selected source', error);
      useLiveFallback();
    } finally {
      await resolvedSource?.cleanup?.().catch(() => undefined);
    }
  };

  const scheduleRefresh = (delayMs = DEFAULT_REFRESH_DELAY_MS): void => {
    if (refreshTimer) clearTimeout(refreshTimer);
    refreshTimer = setTimeout(() => {
      refreshTimer = null;
      void refresh();
    }, delayMs);
  };

  const clearSelectedTrack = (): void => {
    refreshGeneration += 1;
    if (refreshTimer) clearTimeout(refreshTimer);
    refreshTimer = null;
    parsedCues = null;
    parsedSourceKey = null;
    parsedTrackIdentity = null;
    secondaryDelaySeconds = 0;
    lastLiveText = '';
    publish('');
  };

  return {
    refresh,
    scheduleRefresh,
    handleLiveText(text: string): void {
      lastLiveText = text;
      publish(resolveAtTime(deps.getCurrentTimePos()));
    },
    handleTimePos(timeSeconds: number): void {
      if (!parsedCues) return;
      publish(resolveAtTime(timeSeconds));
    },
    handleTrackChange(): void {
      clearSelectedTrack();
    },
    handleDelayChange(delaySeconds: number): void {
      secondaryDelaySeconds = finiteNumber(delaySeconds);
      if (parsedCues) {
        publish(resolveAtTime(deps.getCurrentTimePos()));
      }
    },
    reset: clearSelectedTrack,
  };
}
