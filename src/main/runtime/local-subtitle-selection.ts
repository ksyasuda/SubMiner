import path from 'node:path';

import { isRemoteMediaPath } from '../../jimaku/utils';
import { normalizeYoutubeLangCode } from '../../core/services/youtube/labels';

const DEFAULT_PRIMARY_SUBTITLE_LANGUAGES = ['ja', 'jpn'];
const DEFAULT_SECONDARY_SUBTITLE_LANGUAGES = ['en', 'eng', 'english', 'enus', 'en-us'];
const HEARING_IMPAIRED_PATTERN = /\b(hearing impaired|sdh|closed captions?|cc)\b/i;

type SubtitleTrackLike = {
  type?: unknown;
  id?: unknown;
  lang?: unknown;
  title?: unknown;
  external?: unknown;
  selected?: unknown;
};

type NormalizedSubtitleTrack = {
  id: number;
  lang: string;
  title: string;
  external: boolean;
  selected: boolean;
};

export type ManagedLocalSubtitleSelection = {
  primaryTrackId: number | null;
  secondaryTrackId: number | null;
  hasPrimaryMatch: boolean;
  hasSecondaryMatch: boolean;
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

function normalizeTrack(entry: unknown): NormalizedSubtitleTrack | null {
  if (!entry || typeof entry !== 'object') {
    return null;
  }
  const track = entry as SubtitleTrackLike;
  const id = parseTrackId(track.id);
  if (id === null || (track.type !== undefined && track.type !== 'sub')) {
    return null;
  }
  return {
    id,
    lang: String(track.lang || '').trim(),
    title: String(track.title || '').trim(),
    external: track.external === true,
    selected: track.selected === true,
  };
}

function normalizeLanguageList(values: string[], fallback: string[]): string[] {
  const normalized = values
    .map((value) => normalizeYoutubeLangCode(value))
    .filter((value, index, items) => value.length > 0 && items.indexOf(value) === index);
  if (normalized.length > 0) {
    return normalized;
  }
  return fallback
    .map((value) => normalizeYoutubeLangCode(value))
    .filter((value, index, items) => value.length > 0 && items.indexOf(value) === index);
}

function resolveLanguageRank(language: string, preferredLanguages: string[]): number {
  const normalized = normalizeYoutubeLangCode(language);
  if (!normalized) {
    return Number.POSITIVE_INFINITY;
  }
  const directIndex = preferredLanguages.indexOf(normalized);
  if (directIndex >= 0) {
    return directIndex;
  }
  const base = normalized.split('-')[0] || normalized;
  const baseIndex = preferredLanguages.indexOf(base);
  return baseIndex >= 0 ? baseIndex : Number.POSITIVE_INFINITY;
}

function isLikelyHearingImpaired(title: string): boolean {
  return HEARING_IMPAIRED_PATTERN.test(title);
}

function isUnlabeledExternalTrack(track: NormalizedSubtitleTrack): boolean {
  return track.external && normalizeYoutubeLangCode(track.lang).length === 0;
}

function pickBestTrackId(
  tracks: NormalizedSubtitleTrack[],
  preferredLanguages: string[],
  excludeId: number | null = null,
): { trackId: number | null; hasMatch: boolean } {
  const ranked = tracks
    .filter((track) => track.id !== excludeId)
    .map((track) => ({
      track,
      languageRank: resolveLanguageRank(track.lang, preferredLanguages),
    }))
    .filter(({ languageRank }) => Number.isFinite(languageRank))
    .sort((left, right) => {
      if (left.languageRank !== right.languageRank) {
        return left.languageRank - right.languageRank;
      }
      if (left.track.external !== right.track.external) {
        return left.track.external ? -1 : 1;
      }
      if (
        isLikelyHearingImpaired(left.track.title) !== isLikelyHearingImpaired(right.track.title)
      ) {
        return isLikelyHearingImpaired(left.track.title) ? 1 : -1;
      }
      if (/\bdefault\b/i.test(left.track.title) !== /\bdefault\b/i.test(right.track.title)) {
        return /\bdefault\b/i.test(left.track.title) ? -1 : 1;
      }
      return left.track.id - right.track.id;
    });

  return {
    trackId: ranked[0]?.track.id ?? null,
    hasMatch: ranked.length > 0,
  };
}

function pickSingleUnlabeledExternalTrackId(
  tracks: NormalizedSubtitleTrack[],
  excludeId: number | null = null,
): number | null {
  const fallbackCandidates = tracks.filter(
    (track) => track.id !== excludeId && isUnlabeledExternalTrack(track),
  );
  if (fallbackCandidates.length !== 1) {
    return null;
  }
  return fallbackCandidates[0]?.id ?? null;
}

export function resolveManagedLocalSubtitleSelection(input: {
  trackList: unknown[] | null;
  primaryLanguages: string[];
  secondaryLanguages: string[];
}): ManagedLocalSubtitleSelection {
  const tracks = Array.isArray(input.trackList)
    ? input.trackList
        .map(normalizeTrack)
        .filter((track): track is NormalizedSubtitleTrack => track !== null)
    : [];
  const preferredPrimaryLanguages = normalizeLanguageList(
    input.primaryLanguages,
    DEFAULT_PRIMARY_SUBTITLE_LANGUAGES,
  );
  const preferredSecondaryLanguages = normalizeLanguageList(
    input.secondaryLanguages,
    DEFAULT_SECONDARY_SUBTITLE_LANGUAGES,
  );

  const primary = pickBestTrackId(tracks, preferredPrimaryLanguages);
  const primaryTrackId = primary.trackId ?? pickSingleUnlabeledExternalTrackId(tracks);
  const secondary = pickBestTrackId(tracks, preferredSecondaryLanguages, primaryTrackId);

  return {
    primaryTrackId,
    secondaryTrackId: secondary.trackId,
    hasPrimaryMatch: primary.hasMatch || primaryTrackId !== null,
    hasSecondaryMatch: secondary.hasMatch,
  };
}

function normalizeLocalMediaPath(mediaPath: string | null | undefined): string | null {
  if (typeof mediaPath !== 'string') {
    return null;
  }
  const trimmed = mediaPath.trim();
  if (!trimmed || isRemoteMediaPath(trimmed)) {
    return null;
  }
  return path.resolve(trimmed);
}

export function createManagedLocalSubtitleSelectionRuntime(deps: {
  getCurrentMediaPath: () => string | null;
  getMpvClient: () => {
    connected?: boolean;
    requestProperty?: (name: string) => Promise<unknown>;
  } | null;
  getPrimarySubtitleLanguages: () => string[];
  getSecondarySubtitleLanguages: () => string[];
  sendMpvCommand: (command: ['set_property', 'sid' | 'secondary-sid', number]) => void;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  clearScheduled: (timer: ReturnType<typeof setTimeout>) => void;
  delayMs?: number;
}) {
  const delayMs = deps.delayMs ?? 400;
  let currentMediaPath: string | null = null;
  let appliedMediaPath: string | null = null;
  let pendingTimer: ReturnType<typeof setTimeout> | null = null;

  const clearPendingTimer = (): void => {
    if (!pendingTimer) {
      return;
    }
    deps.clearScheduled(pendingTimer);
    pendingTimer = null;
  };

  const maybeApplySelection = (trackList: unknown[] | null): void => {
    if (!currentMediaPath || appliedMediaPath === currentMediaPath) {
      return;
    }
    const selection = resolveManagedLocalSubtitleSelection({
      trackList,
      primaryLanguages: deps.getPrimarySubtitleLanguages(),
      secondaryLanguages: deps.getSecondarySubtitleLanguages(),
    });
    if (!selection.hasPrimaryMatch && !selection.hasSecondaryMatch) {
      return;
    }
    if (selection.primaryTrackId !== null) {
      deps.sendMpvCommand(['set_property', 'sid', selection.primaryTrackId]);
    }
    if (selection.secondaryTrackId !== null) {
      deps.sendMpvCommand(['set_property', 'secondary-sid', selection.secondaryTrackId]);
    }
    appliedMediaPath = currentMediaPath;
    clearPendingTimer();
  };

  const refreshFromMpv = async (): Promise<void> => {
    const client = deps.getMpvClient();
    if (!client?.connected || !client.requestProperty) {
      return;
    }
    const mediaPath = normalizeLocalMediaPath(deps.getCurrentMediaPath());
    if (!mediaPath || mediaPath !== currentMediaPath) {
      return;
    }
    try {
      const trackList = await client.requestProperty('track-list');
      maybeApplySelection(Array.isArray(trackList) ? trackList : null);
    } catch {
      // Skip selection when mpv track inspection fails.
    }
  };

  const scheduleRefresh = (): void => {
    clearPendingTimer();
    if (!currentMediaPath || appliedMediaPath === currentMediaPath) {
      return;
    }
    pendingTimer = deps.schedule(() => {
      pendingTimer = null;
      void refreshFromMpv();
    }, delayMs);
  };

  return {
    handleMediaPathChange: (mediaPath: string | null | undefined): void => {
      const normalizedPath = normalizeLocalMediaPath(mediaPath);
      if (normalizedPath !== currentMediaPath) {
        appliedMediaPath = null;
      }
      currentMediaPath = normalizedPath;
      if (!currentMediaPath) {
        clearPendingTimer();
        return;
      }
      scheduleRefresh();
    },
    handleSubtitleTrackListChange: (trackList: unknown[] | null): void => {
      maybeApplySelection(trackList);
    },
  };
}
