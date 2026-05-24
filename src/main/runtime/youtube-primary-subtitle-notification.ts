import { isYoutubeMediaPath } from './youtube-playback';
import { normalizeYoutubeLangCode } from '../../core/services/youtube/labels';

export type YoutubePrimarySubtitleNotificationTimer =
  | ReturnType<typeof setTimeout>
  | { id: number };

type SubtitleTrackEntry = {
  id: number | null;
  type: string;
  lang: string;
  external: boolean;
  selected: boolean;
};

type CurrentSubtitleState = {
  sid: unknown;
  trackList: unknown[] | null;
};

function parseTrackId(value: unknown): number | null {
  if (typeof value === 'number' && Number.isInteger(value)) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number(value.trim());
    return Number.isInteger(parsed) ? parsed : null;
  }
  return null;
}

function normalizeTrack(entry: unknown): SubtitleTrackEntry | null {
  if (!entry || typeof entry !== 'object') {
    return null;
  }
  const track = entry as Record<string, unknown>;
  return {
    id: parseTrackId(track.id),
    type: String(track.type || '').trim(),
    lang: String(track.lang || '').trim(),
    external: track.external === true,
    selected: track.selected === true,
  };
}

export function clearYoutubePrimarySubtitleNotificationTimer(
  timer: YoutubePrimarySubtitleNotificationTimer | null,
): void {
  if (!timer) {
    return;
  }
  if (typeof timer === 'object' && timer !== null && 'id' in timer) {
    clearTimeout((timer as { id: number }).id);
    return;
  }
  clearTimeout(timer);
}

function buildPreferredLanguageSet(values: string[]): Set<string> {
  const normalized = values
    .map((value) => normalizeYoutubeLangCode(value))
    .filter((value) => value.length > 0);
  return new Set(normalized);
}

function matchesPreferredLanguage(language: string, preferred: Set<string>): boolean {
  if (preferred.size === 0) {
    return false;
  }
  const normalized = normalizeYoutubeLangCode(language);
  if (!normalized) {
    return false;
  }
  if (preferred.has(normalized)) {
    return true;
  }
  const base = normalized.split('-')[0] || normalized;
  return preferred.has(base);
}

function hasSelectedPrimarySubtitle(
  sid: number | null,
  trackList: unknown[] | null,
  preferredLanguages: Set<string>,
): boolean {
  if (!Array.isArray(trackList)) {
    return false;
  }

  const tracks = trackList.map(normalizeTrack);
  const activeTrack =
    (sid === null
      ? null
      : (tracks.find((track) => track?.type === 'sub' && track.id === sid) ?? null)) ??
    tracks.find((track) => track?.type === 'sub' && track.selected) ??
    null;
  if (!activeTrack) {
    return false;
  }
  if (activeTrack.external) {
    return true;
  }
  return matchesPreferredLanguage(activeTrack.lang, preferredLanguages);
}

export function createYoutubePrimarySubtitleNotificationRuntime(deps: {
  getPrimarySubtitleLanguages: () => string[];
  notifyFailure: (message: string) => void;
  schedule: (
    fn: () => void | Promise<void>,
    delayMs: number,
  ) => YoutubePrimarySubtitleNotificationTimer;
  clearSchedule: (timer: YoutubePrimarySubtitleNotificationTimer | null) => void;
  getCurrentSubtitleState?: () => CurrentSubtitleState | null | Promise<CurrentSubtitleState | null>;
  delayMs?: number;
}) {
  const delayMs = deps.delayMs ?? 5000;
  let currentMediaPath: string | null = null;
  let currentSid: number | null = null;
  let currentTrackList: unknown[] | null = null;
  let pendingTimer: YoutubePrimarySubtitleNotificationTimer | null = null;
  let lastReportedMediaPath: string | null = null;
  let appOwnedFlowInFlight = false;
  let primarySubtitleLoadedForCurrentMedia = false;

  const clearPendingTimer = (): void => {
    deps.clearSchedule(pendingTimer);
    pendingTimer = null;
  };

  const refreshCurrentSubtitleState = async (
    preferredLanguages: Set<string>,
  ): Promise<boolean> => {
    const getCurrentSubtitleState = deps.getCurrentSubtitleState;
    if (!getCurrentSubtitleState) {
      return false;
    }
    let state: CurrentSubtitleState | null;
    try {
      state = await getCurrentSubtitleState();
    } catch {
      state = null;
    }
    if (!state) {
      return false;
    }
    currentSid = parseTrackId(state.sid);
    currentTrackList = Array.isArray(state.trackList) ? state.trackList : null;
    return hasSelectedPrimarySubtitle(currentSid, currentTrackList, preferredLanguages);
  };

  const maybeReportFailure = async (): Promise<void> => {
    const mediaPath = currentMediaPath?.trim() || '';
    if (!mediaPath || !isYoutubeMediaPath(mediaPath)) {
      return;
    }
    if (lastReportedMediaPath === mediaPath) {
      return;
    }
    if (appOwnedFlowInFlight) {
      return;
    }
    const preferredLanguages = buildPreferredLanguageSet(deps.getPrimarySubtitleLanguages());
    if (preferredLanguages.size === 0) {
      return;
    }
    if (primarySubtitleLoadedForCurrentMedia) {
      return;
    }
    if (hasSelectedPrimarySubtitle(currentSid, currentTrackList, preferredLanguages)) {
      return;
    }
    if (deps.getCurrentSubtitleState && (await refreshCurrentSubtitleState(preferredLanguages))) {
      clearPendingTimer();
      return;
    }
    if (
      currentMediaPath?.trim() !== mediaPath ||
      appOwnedFlowInFlight ||
      primarySubtitleLoadedForCurrentMedia
    ) {
      return;
    }
    lastReportedMediaPath = mediaPath;
    deps.notifyFailure(
      'Primary subtitle failed to download or load. Try again from the subtitle modal.',
    );
  };

  const schedulePendingCheck = (): void => {
    clearPendingTimer();
    if (appOwnedFlowInFlight) {
      return;
    }
    const mediaPath = currentMediaPath?.trim() || '';
    if (!mediaPath || !isYoutubeMediaPath(mediaPath)) {
      return;
    }
    if (primarySubtitleLoadedForCurrentMedia) {
      return;
    }
    pendingTimer = deps.schedule(async () => {
      pendingTimer = null;
      await maybeReportFailure();
    }, delayMs);
  };

  return {
    handleMediaPathChange: (path: string | null): void => {
      const normalizedPath =
        typeof path === 'string' && path.trim().length > 0 ? path.trim() : null;
      if (currentMediaPath !== normalizedPath) {
        lastReportedMediaPath = null;
        primarySubtitleLoadedForCurrentMedia = false;
      }
      currentMediaPath = normalizedPath;
      currentSid = null;
      currentTrackList = null;
      schedulePendingCheck();
    },
    handleSubtitleTrackChange: (sid: number | null): void => {
      currentSid = sid;
      const preferredLanguages = buildPreferredLanguageSet(deps.getPrimarySubtitleLanguages());
      if (hasSelectedPrimarySubtitle(currentSid, currentTrackList, preferredLanguages)) {
        clearPendingTimer();
      }
    },
    handleSubtitleTrackListChange: (trackList: unknown[] | null): void => {
      currentTrackList = trackList;
      const preferredLanguages = buildPreferredLanguageSet(deps.getPrimarySubtitleLanguages());
      if (hasSelectedPrimarySubtitle(currentSid, currentTrackList, preferredLanguages)) {
        clearPendingTimer();
      }
    },
    markCurrentMediaPrimarySubtitleLoaded: (): void => {
      const mediaPath = currentMediaPath?.trim() || '';
      if (!mediaPath || !isYoutubeMediaPath(mediaPath)) {
        return;
      }
      primarySubtitleLoadedForCurrentMedia = true;
      clearPendingTimer();
    },
    setAppOwnedFlowInFlight: (inFlight: boolean): void => {
      appOwnedFlowInFlight = inFlight;
      if (inFlight) {
        clearPendingTimer();
        return;
      }
      schedulePendingCheck();
    },
    isAppOwnedFlowInFlight: (): boolean => appOwnedFlowInFlight,
  };
}
