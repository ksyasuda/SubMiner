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
  schedule: (fn: () => void, delayMs: number) => YoutubePrimarySubtitleNotificationTimer;
  clearSchedule: (timer: YoutubePrimarySubtitleNotificationTimer | null) => void;
  delayMs?: number;
}) {
  const delayMs = deps.delayMs ?? 5000;
  let currentMediaPath: string | null = null;
  let currentSid: number | null = null;
  let currentTrackList: unknown[] | null = null;
  let pendingTimer: YoutubePrimarySubtitleNotificationTimer | null = null;
  let lastReportedMediaPath: string | null = null;
  let appOwnedFlowInFlight = false;

  const clearPendingTimer = (): void => {
    deps.clearSchedule(pendingTimer);
    pendingTimer = null;
  };

  const maybeReportFailure = (): void => {
    const mediaPath = currentMediaPath?.trim() || '';
    if (!mediaPath || !isYoutubeMediaPath(mediaPath)) {
      return;
    }
    if (lastReportedMediaPath === mediaPath) {
      return;
    }
    const preferredLanguages = buildPreferredLanguageSet(deps.getPrimarySubtitleLanguages());
    if (preferredLanguages.size === 0) {
      return;
    }
    if (hasSelectedPrimarySubtitle(currentSid, currentTrackList, preferredLanguages)) {
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
    pendingTimer = deps.schedule(() => {
      pendingTimer = null;
      maybeReportFailure();
    }, delayMs);
  };

  return {
    handleMediaPathChange: (path: string | null): void => {
      const normalizedPath =
        typeof path === 'string' && path.trim().length > 0 ? path.trim() : null;
      if (currentMediaPath !== normalizedPath) {
        lastReportedMediaPath = null;
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
