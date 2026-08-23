import type { SubtitleCue } from '../../types/subtitle';
import { flattenedSecondarySubtitleLineIdentity } from '../../core/services/secondary-subtitle-line-identity';
import { removeAssControlDebrisLines } from '../../core/services/ass-text';

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

function sourceUsesAssSyntax(source: string): boolean {
  const sourceWithoutQuery = source.split(/[?#]/u, 1)[0] ?? '';
  return /\.(?:ass|ssa)$/iu.test(sourceWithoutQuery);
}

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

type IndexedSubtitleCue = { cue: SubtitleCue; index: number };

function compareAuthoredSubtitleOrder(left: IndexedSubtitleCue, right: IndexedSubtitleCue): number {
  const leftLayout = left.cue.assLayout;
  const rightLayout = right.cue.assLayout;
  if (leftLayout?.kind === 'positioned' && rightLayout?.kind === 'positioned') {
    const verticalOrder = leftLayout.y - rightLayout.y;
    if (verticalOrder !== 0) return verticalOrder;
  }
  if (leftLayout && rightLayout) {
    const sourceOrder = leftLayout.sourceOrder - rightLayout.sourceOrder;
    if (sourceOrder !== 0) return sourceOrder;
  }
  return left.index - right.index;
}

export function findActiveSubtitleText(cues: readonly SubtitleCue[], timeSeconds: number): string {
  if (!Number.isFinite(timeSeconds)) return '';

  const authoredCanonical = cues.filter(
    (cue) =>
      cue.source === 'canonical-ass' && cue.startTime <= timeSeconds && cue.endTime > timeSeconds,
  );
  const selectedCanonical = new Set<SubtitleCue>(authoredCanonical);
  if (selectedCanonical.size === 0) {
    const animatedCanonical = cues.filter(
      (cue) =>
        cue.source === 'canonical-ass' &&
        (cue.animationStartTime ?? cue.startTime) <= timeSeconds &&
        (cue.animationEndTime ?? cue.endTime) > timeSeconds,
    );
    const nearestDistance = animatedCanonical.reduce((nearest, cue) => {
      const distance =
        timeSeconds < cue.startTime
          ? cue.startTime - timeSeconds
          : Math.max(0, timeSeconds - cue.endTime);
      return Math.min(nearest, distance);
    }, Infinity);
    for (const cue of animatedCanonical) {
      const distance =
        timeSeconds < cue.startTime
          ? cue.startTime - timeSeconds
          : Math.max(0, timeSeconds - cue.endTime);
      if (distance === nearestDistance) {
        selectedCanonical.add(cue);
      }
    }
  }

  const activeReconstructed = cues.filter(
    (cue) =>
      cue.source === 'reconstructed-ass' &&
      cue.assLayout?.kind !== 'fragment-grid' &&
      cue.startTime <= timeSeconds &&
      cue.endTime > timeSeconds,
  );
  const reconstructedByStyle = new Map<string, SubtitleCue>();
  for (const cue of activeReconstructed) {
    const style = cue.assStyle ?? '';
    const existing = reconstructedByStyle.get(style);
    if (!existing) {
      reconstructedByStyle.set(style, cue);
      continue;
    }
    const duration = cue.endTime - cue.startTime;
    const existingDuration = existing.endTime - existing.startTime;
    if (
      duration > existingDuration ||
      (duration === existingDuration && cue.text.length > existing.text.length) ||
      (duration === existingDuration &&
        cue.text.length === existing.text.length &&
        cue.startTime > existing.startTime)
    ) {
      reconstructedByStyle.set(style, cue);
    }
  }
  const selectedReconstructed = new Set(reconstructedByStyle.values());

  const seenExact = new Set<string>();
  const seenFlattened = new Set<string>();
  const activeText: string[] = [];
  const activeCues: IndexedSubtitleCue[] = [];
  cues.forEach((cue, index) => {
    const active =
      cue.source === 'canonical-ass'
        ? selectedCanonical.has(cue)
        : cue.source === 'reconstructed-ass'
          ? selectedReconstructed.has(cue)
          : cue.startTime <= timeSeconds && cue.endTime > timeSeconds;
    if (active) activeCues.push({ cue, index });
  });
  activeCues.sort(compareAuthoredSubtitleOrder);

  for (const { cue } of activeCues) {
    for (const line of cue.text.split('\n')) {
      const text = line.trim();
      const compactText = text.normalize('NFKC').replace(/\s+/gu, '');
      if (!compactText || seenExact.has(compactText)) continue;
      seenExact.add(compactText);

      const flattenedIdentity = flattenedSecondarySubtitleLineIdentity(text);
      if (flattenedIdentity && seenFlattened.has(flattenedIdentity)) continue;
      if (flattenedIdentity) seenFlattened.add(flattenedIdentity);
      activeText.push(text);
    }
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
  let activeSourceUsesAssSyntax = false;
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
        activeSourceUsesAssSyntax = false;
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
        activeSourceUsesAssSyntax = false;
        deps.logDebug?.('[secondary-subtitle-track] selected source is not readable');
        useLiveFallback();
        return;
      }

      activeSourceUsesAssSyntax = sourceUsesAssSyntax(resolvedSource.path);

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
    activeSourceUsesAssSyntax = false;
    secondaryDelaySeconds = 0;
    lastLiveText = '';
    publish('');
  };

  return {
    refresh,
    scheduleRefresh,
    handleLiveText(text: string): void {
      lastLiveText = activeSourceUsesAssSyntax ? removeAssControlDebrisLines(text) : text;
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
