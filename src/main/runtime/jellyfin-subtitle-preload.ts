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
    .flatMap(({ track, cached }) =>
      cached
        ? [
            {
              track,
              score:
                (track.external ? 100 : 0) +
                scoreJellyfinSource(cached.source, cached.source.title || track.title),
            },
          ]
        : [],
    )
    .sort((a, b) => b.score - a.score);
  return ranked[0]?.track.id ?? null;
}

// Ranks Jellyfin subtitle sources by metadata alone, so the preferred track is known before download.
function scoreJellyfinSource(source: JellyfinSubtitleTrack, title: string): number {
  return (
    (source.isDefault ? 35 : 0) +
    (source.isExternal === false ? 25 : 0) +
    (source.isExternal === true ? -10 : 0) +
    (source.isForced ? -25 : 0) +
    (isLikelyHearingImpaired(title) ? -10 : 10) +
    (/\bdefault\b/i.test(title) ? 3 : 0)
  );
}

function pickPreferredJapaneseSource(
  sources: JellyfinSubtitleTrack[],
): JellyfinSubtitleTrack | null {
  const ranked = sources
    .filter((source) => isJapanese(source.language || '') || isJapanese(source.title || ''))
    .map((source) => ({ source, score: scoreJellyfinSource(source, source.title || '') }))
    .sort((a, b) => b.score - a.score);
  return ranked[0]?.source ?? null;
}

function findMpvTrackIdByPath(tracks: MpvSubtitleTrack[], filePath: string): number | null {
  return tracks.find((track) => track.externalFilename === filePath)?.id ?? null;
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

  function selectJapanesePrimary(
    subtitleTracks: MpvSubtitleTrack[],
    cachedTracks: CachedExternalSubtitleTrack[],
    trackId: number | null,
  ): void {
    if (trackId === null) {
      deps.sendMpvCommand(['set_property', 'sid', 'no']);
      return;
    }
    deps.sendMpvCommand(['set_property', 'sid', trackId]);
    const selectedCachedTrack = findCachedTrackForMpvTrackId(subtitleTracks, cachedTracks, trackId);
    if (selectedCachedTrack) {
      startSubtitlePrefetchForCachedTrack(selectedCachedTrack.path);
    }
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
      resetManagedSubtitleDelay();
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
        return;
      }

      deps.sendMpvCommand(['set_property', 'sid', 'no']);
      deps.sendMpvCommand(['set_property', 'secondary-sid', 'no']);
      deps.sendMpvCommand(['set_property', 'sub-visibility', 'no']);
      deps.sendMpvCommand(['set_property', 'secondary-sub-visibility', 'no']);
      await deps.wait(300);
      const seenUrls = new Set<string>();
      const uniqueTracks = externalTracks.filter((track) => {
        if (!track.deliveryUrl || seenUrls.has(track.deliveryUrl)) return false;
        seenUrls.add(track.deliveryUrl);
        return true;
      });

      // Download every track at once and add each to mpv as soon as it lands. Jellyfin has
      // to extract embedded tracks from the container, so one slow track must not hold up
      // the Japanese primary that annotations depend on.
      const cachedTracks: CachedExternalSubtitleTrack[] = [];
      const downloads = new Map(
        uniqueTracks.map((track) => [
          track,
          (async (): Promise<CachedExternalSubtitleTrack> => {
            const labelBase = (track.title || track.language || '').trim();
            const label = labelBase || `Jellyfin Subtitle ${track.index}`;
            const cached = { ...(await deps.cacheSubtitleTrack(track)), source: track };
            activeCacheDirs.add(cached.cleanupDir);
            cachedTracks.push(cached);
            deps.sendMpvCommand(['sub-add', cached.path, 'auto', label, track.language || '']);
            return cached;
          })(),
        ]),
      );
      // Settled up front so a failed download is not flagged as unhandled while the
      // Japanese track is still being selected.
      const allDownloads = Promise.allSettled(downloads.values());

      try {
        let subtitleTracks: MpvSubtitleTrack[] = [];
        let japanesePrimaryId: number | null | undefined;

        const preferredJapaneseSource = pickPreferredJapaneseSource(uniqueTracks);
        const preferredJapaneseDownload = preferredJapaneseSource
          ? downloads.get(preferredJapaneseSource)
          : undefined;
        if (preferredJapaneseDownload) {
          const preferredJapanese = await preferredJapaneseDownload;
          await deps.wait(TRACK_SELECTION_INITIAL_WAIT_MS);
          subtitleTracks =
            (await waitForPreferredSubtitleTracks(deps, true, [preferredJapanese.path])) ?? [];
          if (!hasExternalJapaneseTrack(subtitleTracks)) {
            deps.logDebug('Timed out waiting for Jellyfin Japanese subtitle track', {
              itemId: params.itemId,
            });
            return;
          }
          japanesePrimaryId =
            findMpvTrackIdByPath(subtitleTracks, preferredJapanese.path) ??
            pickBestTrackId(subtitleTracks, isJapanese);
          selectJapanesePrimary(subtitleTracks, cachedTracks, japanesePrimaryId);
        }

        const results = await allDownloads;
        const failedDownload = results.find((result) => result.status === 'rejected');
        if (failedDownload) {
          throw failedDownload.reason;
        }
        const cachedPaths = cachedTracks.map((track) => track.path);
        if (!hasExpectedExternalSubtitleTracks(subtitleTracks, cachedPaths)) {
          await deps.wait(TRACK_SELECTION_INITIAL_WAIT_MS);
          subtitleTracks = (await waitForPreferredSubtitleTracks(deps, false, cachedPaths)) ?? [];
        }

        if (japanesePrimaryId === undefined) {
          japanesePrimaryId =
            pickBestCachedTrackId(subtitleTracks, cachedTracks, isJapanese) ??
            pickBestTrackId(subtitleTracks, isJapanese);
          selectJapanesePrimary(subtitleTracks, cachedTracks, japanesePrimaryId);
        }

        const englishSecondaryId =
          pickBestCachedTrackId(subtitleTracks, cachedTracks, isEnglish, japanesePrimaryId) ??
          pickBestTrackId(subtitleTracks, isEnglish, japanesePrimaryId);
        if (englishSecondaryId !== null) {
          deps.sendMpvCommand(['set_property', 'secondary-sid', englishSecondaryId]);
        }
      } finally {
        // Keep this run in the queue until every download has registered its cache dir, so
        // the next run's cleanup sees them all.
        await allDownloads;
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
