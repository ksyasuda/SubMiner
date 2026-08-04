import { parseSubtitleCues } from '../../core/services/subtitle-cue-parser';
import { estimateSubtitleTimingOffset } from '../../core/services/subtitle-timing-offset';

type JellyfinSession = {
  serverUrl: string;
  accessToken: string;
  userId: string;
  username: string;
};

type JellyfinClientInfo = {
  clientName: string;
  clientVersion: string;
  deviceId: string;
};

type JellyfinSubtitleTrack = {
  index: number;
  language?: string;
  title?: string;
  codec?: string;
  isDefault?: boolean;
  isForced?: boolean;
  isExternal?: boolean;
  deliveryMethod?: string;
  deliveryUrl?: string | null;
};

type CachedSubtitleTrack = {
  path: string;
  cleanupDir: string;
};

type CachedExternalSubtitleTrack = CachedSubtitleTrack & {
  source: JellyfinSubtitleTrack;
};

type JellyfinSubtitleDelayKey = {
  itemId: string;
  streamIndex: number;
};

type MpvSubtitleTrack = {
  id: number;
  lang: string;
  title: string;
  external: boolean;
  externalFilename: string;
};

type MpvClientLike = {
  connected?: boolean;
  requestProperty: (name: string) => Promise<unknown>;
};

const TRACK_SELECTION_INITIAL_WAIT_MS = 250;
const TRACK_SELECTION_RETRY_MS = 150;
const TRACK_SELECTION_MAX_ATTEMPTS = 10;

export type PreloadJellyfinExternalSubtitlesHandler = ((params: {
  session: JellyfinSession;
  clientInfo: JellyfinClientInfo;
  itemId: string;
}) => Promise<void>) & {
  cleanupCachedSubtitles: () => void;
};

function normalizeLang(value: unknown): string {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/_/g, '-');
}

function isJapanese(value: string): boolean {
  const v = normalizeLang(value);
  return (
    v === 'ja' ||
    v === 'jp' ||
    v === 'jpn' ||
    v === 'japanese' ||
    v.startsWith('ja-') ||
    v.startsWith('jp-')
  );
}

function isEnglish(value: string): boolean {
  const v = normalizeLang(value);
  return (
    v === 'en' ||
    v === 'eng' ||
    v === 'english' ||
    v === 'enus' ||
    v === 'en-us' ||
    v.startsWith('en-')
  );
}

function isLikelyHearingImpaired(title: string): boolean {
  return /\b(hearing impaired|sdh|closed captions?|cc)\b/i.test(title);
}

function pickBestTrackId(
  tracks: MpvSubtitleTrack[],
  languageMatcher: (value: string) => boolean,
  excludeId: number | null = null,
): number | null {
  const ranked = tracks
    .filter((track) => languageMatcher(track.lang) || languageMatcher(track.title))
    .filter((track) => track.id !== excludeId)
    .map((track) => ({
      track,
      score:
        (track.external ? 100 : 0) +
        (isLikelyHearingImpaired(track.title) ? -10 : 10) +
        (/\bdefault\b/i.test(track.title) ? 3 : 0),
    }))
    .sort((a, b) => b.score - a.score);
  return ranked[0]?.track.id ?? null;
}

function pickBestCachedTrackId(
  tracks: MpvSubtitleTrack[],
  cachedTracks: CachedExternalSubtitleTrack[],
  sourceMatcher: (value: string) => boolean,
  excludeId: number | null = null,
): number | null {
  const cachedByPath = new Map(cachedTracks.map((track) => [track.path, track]));
  const ranked = tracks
    .map((track) => ({
      track,
      cached: cachedByPath.get(track.externalFilename),
    }))
    .filter(({ cached }) =>
      cached
        ? sourceMatcher(cached.source.language || '') || sourceMatcher(cached.source.title || '')
        : false,
    )
    .filter(({ track }) => track.id !== excludeId)
    .map(({ track, cached }) => {
      const title = cached?.source.title || track.title;
      return {
        track,
        score:
          (track.external ? 100 : 0) +
          (cached?.source.isDefault ? 35 : 0) +
          (cached?.source.isExternal === false ? 25 : 0) +
          (cached?.source.isExternal === true ? -10 : 0) +
          (cached?.source.isForced ? -25 : 0) +
          (isLikelyHearingImpaired(title) ? -10 : 10) +
          (/\bdefault\b/i.test(title) ? 3 : 0),
      };
    })
    .sort((a, b) => b.score - a.score);
  return ranked[0]?.track.id ?? null;
}

function findCachedTrackForMpvTrackId(
  tracks: MpvSubtitleTrack[],
  cachedTracks: CachedExternalSubtitleTrack[],
  trackId: number | null,
): CachedExternalSubtitleTrack | null {
  if (trackId === null) return null;
  const mpvTrack = tracks.find((track) => track.id === trackId);
  if (!mpvTrack?.externalFilename) return null;
  return cachedTracks.find((track) => track.path === mpvTrack.externalFilename) ?? null;
}

function isJapaneseTrack(track: MpvSubtitleTrack): boolean {
  return isJapanese(track.lang) || isJapanese(track.title);
}

function hasExternalJapaneseTrack(tracks: MpvSubtitleTrack[]): boolean {
  return tracks.some((track) => track.external && isJapaneseTrack(track));
}

function parseMpvSubtitleTracks(trackListRaw: unknown): MpvSubtitleTrack[] {
  return Array.isArray(trackListRaw)
    ? trackListRaw
        .filter(
          (track): track is Record<string, unknown> =>
            Boolean(track) && typeof track === 'object' && track.type === 'sub',
        )
        .map((track) => ({
          id: parseTrackId(track.id),
          lang: String(track.lang || ''),
          title: String(track.title || ''),
          external: track.external === true,
          externalFilename: String(track['external-filename'] || ''),
        }))
        .filter((track): track is MpvSubtitleTrack => track.id !== null)
    : [];
}

function hasExpectedExternalSubtitleTracks(
  tracks: MpvSubtitleTrack[],
  expectedExternalFilenames: string[],
): boolean {
  if (expectedExternalFilenames.length === 0) {
    return true;
  }
  const loadedExternalFilenames = new Set(
    tracks.filter((track) => track.externalFilename).map((track) => track.externalFilename),
  );
  return expectedExternalFilenames.every((filePath) => loadedExternalFilenames.has(filePath));
}

function parseTrackId(value: unknown): number | null {
  if (typeof value === 'string' && value.trim() === '') {
    return null;
  }
  const numeric =
    typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : NaN;
  return Number.isFinite(numeric) ? numeric : null;
}

async function readMpvSubtitleTracks(deps: {
  getMpvClient: () => MpvClientLike | null;
}): Promise<MpvSubtitleTrack[] | null> {
  const client = deps.getMpvClient();
  if (!client || client.connected === false) {
    return null;
  }
  let trackListRaw: unknown;
  try {
    trackListRaw = await client.requestProperty('track-list');
  } catch {
    return null;
  }
  return parseMpvSubtitleTracks(trackListRaw);
}

async function waitForPreferredSubtitleTracks(
  deps: {
    getMpvClient: () => MpvClientLike | null;
    wait: (ms: number) => Promise<void>;
  },
  shouldWaitForExternalJapanese: boolean,
  expectedExternalFilenames: string[],
): Promise<MpvSubtitleTrack[] | null> {
  let subtitleTracks: MpvSubtitleTrack[] = [];
  for (let attempt = 1; attempt <= TRACK_SELECTION_MAX_ATTEMPTS; attempt += 1) {
    const nextTracks = await readMpvSubtitleTracks(deps);
    if (nextTracks !== null) {
      subtitleTracks = nextTracks;
      if (
        (!shouldWaitForExternalJapanese || hasExternalJapaneseTrack(subtitleTracks)) &&
        hasExpectedExternalSubtitleTracks(subtitleTracks, expectedExternalFilenames)
      ) {
        return subtitleTracks;
      }
    }
    if (attempt < TRACK_SELECTION_MAX_ATTEMPTS) {
      await deps.wait(TRACK_SELECTION_RETRY_MS);
    }
  }
  return subtitleTracks;
}

async function estimateSubtitleDelayFromReference(
  deps: {
    loadSubtitleSourceText?: (source: string) => Promise<string>;
    logDebug: (message: string, error: unknown) => void;
  },
  primaryTrack: CachedExternalSubtitleTrack | null,
  referenceTrack: CachedExternalSubtitleTrack | null,
): Promise<number | null> {
  if (!deps.loadSubtitleSourceText || !primaryTrack || !referenceTrack) {
    return null;
  }

  try {
    const [primaryContent, referenceContent] = await Promise.all([
      deps.loadSubtitleSourceText(primaryTrack.path),
      deps.loadSubtitleSourceText(referenceTrack.path),
    ]);
    const primaryCues = parseSubtitleCues(primaryContent, primaryTrack.path);
    const referenceCues = parseSubtitleCues(referenceContent, referenceTrack.path);
    return estimateSubtitleTimingOffset(primaryCues, referenceCues)?.offsetSeconds ?? null;
  } catch (error) {
    deps.logDebug('Failed to auto-align Jellyfin subtitle timing', error);
    return null;
  }
}

function saveEstimatedSubtitleDelay(
  deps: {
    saveSubtitleDelay?: (
      itemId: string,
      streamIndex: number,
      delaySeconds: number,
    ) => boolean | void;
    logDebug: (message: string, error: unknown) => void;
  },
  key: JellyfinSubtitleDelayKey,
  delaySeconds: number,
): void {
  try {
    const saved = deps.saveSubtitleDelay?.(key.itemId, key.streamIndex, delaySeconds);
    if (saved === false) {
      deps.logDebug('Failed to save Jellyfin auto subtitle delay', key);
    }
  } catch (error) {
    deps.logDebug('Failed to save Jellyfin auto subtitle delay', error);
  }
}

export function createPreloadJellyfinExternalSubtitlesHandler(deps: {
  listJellyfinSubtitleTracks: (
    session: JellyfinSession,
    clientInfo: JellyfinClientInfo,
    itemId: string,
  ) => Promise<JellyfinSubtitleTrack[]>;
  getMpvClient: () => MpvClientLike | null;
  sendMpvCommand: (command: Array<string | number>) => void;
  wait: (ms: number) => Promise<void>;
  cacheSubtitleTrack: (track: JellyfinSubtitleTrack) => Promise<CachedSubtitleTrack>;
  cleanupCachedSubtitles: (dirs: string[]) => void;
  getSavedSubtitleDelay?: (itemId: string, streamIndex: number) => number | null;
  setActiveSubtitleDelayKey?: (key: JellyfinSubtitleDelayKey | null) => void;
  loadSubtitleSourceText?: (source: string) => Promise<string>;
  saveSubtitleDelay?: (itemId: string, streamIndex: number, delaySeconds: number) => boolean | void;
  initSubtitlePrefetch?: (sourcePath: string) => void | Promise<void>;
  logDebug: (message: string, error: unknown) => void;
}): PreloadJellyfinExternalSubtitlesHandler {
  const activeCacheDirs = new Set<string>();
  let preloadQueue: Promise<void> = Promise.resolve();

  function resetManagedSubtitleDelay(): void {
    deps.sendMpvCommand(['set_property', 'sub-delay', 0]);
  }

  // mpv's sid property-change is the only thing that normally starts prefetching, so a
  // coalesced or missed event leaves the whole episode uncached. The downloaded path is
  // known here, so seed the pipeline directly instead of waiting on the observer.
  function startSubtitlePrefetchForCachedTrack(sourcePath: string): void {
    if (!deps.initSubtitlePrefetch) return;
    void Promise.resolve()
      .then(() => deps.initSubtitlePrefetch!(sourcePath))
      .catch((error) => {
        deps.logDebug('Failed to start subtitle prefetch for Jellyfin subtitle', error);
      });
  }

  function cleanupActiveCache(): void {
    const dirs = [...activeCacheDirs];
    if (dirs.length === 0) return;
    deps.cleanupCachedSubtitles(dirs);
    for (const dir of dirs) {
      activeCacheDirs.delete(dir);
    }
  }

  const runPreload = async (params: {
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    itemId: string;
  }): Promise<void> => {
    try {
      try {
        cleanupActiveCache();
      } catch (error) {
        deps.logDebug('Failed to cleanup Jellyfin cached subtitles', error);
      }
      const tracks = await deps.listJellyfinSubtitleTracks(
        params.session,
        params.clientInfo,
        params.itemId,
      );
      const externalTracks = tracks.filter((track) => Boolean(track.deliveryUrl));
      if (externalTracks.length === 0) {
        deps.setActiveSubtitleDelayKey?.(null);
        resetManagedSubtitleDelay();
        return;
      }

      deps.sendMpvCommand(['set_property', 'sid', 'no']);
      deps.sendMpvCommand(['set_property', 'secondary-sid', 'no']);
      deps.sendMpvCommand(['set_property', 'sub-visibility', 'no']);
      deps.sendMpvCommand(['set_property', 'secondary-sub-visibility', 'no']);
      await deps.wait(300);
      const seenUrls = new Set<string>();
      const cachedTracks: CachedExternalSubtitleTrack[] = [];
      for (const track of externalTracks) {
        if (!track.deliveryUrl || seenUrls.has(track.deliveryUrl)) {
          continue;
        }
        seenUrls.add(track.deliveryUrl);
        const labelBase = (track.title || track.language || '').trim();
        const label = labelBase || `Jellyfin Subtitle ${track.index}`;
        const cached = await deps.cacheSubtitleTrack(track);
        activeCacheDirs.add(cached.cleanupDir);
        cachedTracks.push({ ...cached, source: track });
        deps.sendMpvCommand(['sub-add', cached.path, 'auto', label, track.language || '']);
      }

      await deps.wait(TRACK_SELECTION_INITIAL_WAIT_MS);
      const shouldWaitForExternalJapanese = externalTracks.some(
        (track) => isJapanese(track.language || '') || isJapanese(track.title || ''),
      );
      const subtitleTracks = await waitForPreferredSubtitleTracks(
        deps,
        shouldWaitForExternalJapanese,
        cachedTracks.map((track) => track.path),
      );
      if (
        shouldWaitForExternalJapanese &&
        (!subtitleTracks || !hasExternalJapaneseTrack(subtitleTracks))
      ) {
        deps.logDebug('Timed out waiting for Jellyfin Japanese subtitle track', {
          itemId: params.itemId,
        });
        return;
      }

      const resolvedSubtitleTracks = subtitleTracks ?? [];
      const japanesePrimaryId =
        pickBestCachedTrackId(resolvedSubtitleTracks, cachedTracks, isJapanese) ??
        pickBestTrackId(resolvedSubtitleTracks, isJapanese);
      const englishSecondaryId =
        pickBestCachedTrackId(resolvedSubtitleTracks, cachedTracks, isEnglish, japanesePrimaryId) ??
        pickBestTrackId(resolvedSubtitleTracks, isEnglish, japanesePrimaryId);
      if (japanesePrimaryId !== null) {
        const selectedCachedTrack = findCachedTrackForMpvTrackId(
          resolvedSubtitleTracks,
          cachedTracks,
          japanesePrimaryId,
        );
        if (selectedCachedTrack) {
          const delayKey = { itemId: params.itemId, streamIndex: selectedCachedTrack.source.index };
          deps.setActiveSubtitleDelayKey?.(delayKey);
          const savedDelay = deps.getSavedSubtitleDelay?.(delayKey.itemId, delayKey.streamIndex);
          if (typeof savedDelay === 'number' && Number.isFinite(savedDelay)) {
            deps.sendMpvCommand(['set_property', 'sub-delay', savedDelay]);
          } else {
            const referenceCachedTrack = findCachedTrackForMpvTrackId(
              resolvedSubtitleTracks,
              cachedTracks,
              englishSecondaryId,
            );
            const estimatedDelay = await estimateSubtitleDelayFromReference(
              deps,
              selectedCachedTrack,
              referenceCachedTrack,
            );
            if (estimatedDelay !== null) {
              deps.sendMpvCommand(['set_property', 'sub-delay', estimatedDelay]);
              saveEstimatedSubtitleDelay(deps, delayKey, estimatedDelay);
            } else {
              resetManagedSubtitleDelay();
            }
          }
          deps.sendMpvCommand(['set_property', 'sid', japanesePrimaryId]);
          startSubtitlePrefetchForCachedTrack(selectedCachedTrack.path);
        } else {
          deps.setActiveSubtitleDelayKey?.(null);
          resetManagedSubtitleDelay();
          deps.sendMpvCommand(['set_property', 'sid', japanesePrimaryId]);
        }
      } else {
        deps.sendMpvCommand(['set_property', 'sid', 'no']);
        deps.setActiveSubtitleDelayKey?.(null);
        resetManagedSubtitleDelay();
      }

      if (englishSecondaryId !== null) {
        deps.sendMpvCommand(['set_property', 'secondary-sid', englishSecondaryId]);
      }
    } catch (error) {
      deps.logDebug('Failed to preload Jellyfin external subtitles', error);
    }
  };

  const preload = (params: {
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    itemId: string;
  }): Promise<void> => {
    preloadQueue = preloadQueue.then(
      () => runPreload(params),
      () => runPreload(params),
    );
    return preloadQueue;
  };

  return Object.assign(preload, {
    cleanupCachedSubtitles: cleanupActiveCache,
  });
}
