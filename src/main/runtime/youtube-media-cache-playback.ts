import type { YoutubeMediaCacheMode } from '../../types/integrations';
import { isYoutubeMediaPath } from './youtube-playback';

type PlaylistEntry = {
  filename: string;
  current: boolean;
  playing: boolean;
};

export interface YoutubeMediaCachePlaybackRuntimeDeps {
  getMediaCacheConfig: () => {
    mode: YoutubeMediaCacheMode;
    maxHeight?: number;
  };
  requestMpvProperty?: (name: string) => Promise<unknown>;
  startYoutubeMediaCache: (
    url: string,
    options: {
      mode: YoutubeMediaCacheMode;
      maxHeight?: number;
    },
  ) => void;
  logWarn: (message: string) => void;
}

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function normalizePlaylistIndex(value: unknown): number | null {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0 ? value : null;
}

function isGoogleVideoStreamPath(mediaPath: string): boolean {
  try {
    const parsed = new URL(mediaPath);
    const host = parsed.hostname.toLowerCase();
    return host === 'googlevideo.com' || host.endsWith('.googlevideo.com');
  } catch {
    return false;
  }
}

function normalizePlaylistEntries(raw: unknown): PlaylistEntry[] {
  if (!Array.isArray(raw)) {
    return [];
  }
  return raw.map((entry) => {
    const item = (entry ?? {}) as {
      filename?: unknown;
      current?: unknown;
      playing?: unknown;
    };
    return {
      filename: trimToNonEmptyString(item.filename) ?? '',
      current: item.current === true,
      playing: item.playing === true,
    };
  });
}

function resolvePlaylistEntry(
  playlist: PlaylistEntry[],
  playingPosValue: unknown,
): PlaylistEntry | null {
  if (playlist.length === 0) {
    return null;
  }

  const playingPos = normalizePlaylistIndex(playingPosValue);
  if (playingPos !== null && playingPos < playlist.length) {
    return playlist[playingPos] ?? null;
  }

  return playlist.find((entry) => entry.current || entry.playing) ?? null;
}

async function requestMpvPropertySafely(
  deps: YoutubeMediaCachePlaybackRuntimeDeps,
  name: string,
): Promise<unknown> {
  if (!deps.requestMpvProperty) {
    return null;
  }
  try {
    return await deps.requestMpvProperty(name);
  } catch {
    return null;
  }
}

async function resolveYoutubeSourceFromPlaylist(
  deps: YoutubeMediaCachePlaybackRuntimeDeps,
): Promise<string | null> {
  const [playingPosValue, playlistValue] = await Promise.all([
    requestMpvPropertySafely(deps, 'playlist-playing-pos'),
    requestMpvPropertySafely(deps, 'playlist'),
  ]);
  const playlistEntry = resolvePlaylistEntry(
    normalizePlaylistEntries(playlistValue),
    playingPosValue,
  );
  return playlistEntry?.filename && isYoutubeMediaPath(playlistEntry.filename)
    ? playlistEntry.filename
    : null;
}

async function resolveYoutubeSourceUrl(
  deps: YoutubeMediaCachePlaybackRuntimeDeps,
  mediaPath: string,
): Promise<string | null> {
  const directPath = trimToNonEmptyString(mediaPath);
  if (directPath && isYoutubeMediaPath(directPath)) {
    return directPath;
  }
  if (!directPath || !isGoogleVideoStreamPath(directPath)) {
    return null;
  }
  return await resolveYoutubeSourceFromPlaylist(deps);
}

export function createYoutubeMediaCachePlaybackRuntime(deps: YoutubeMediaCachePlaybackRuntimeDeps) {
  let activeYoutubeSourceUrl: string | null = null;
  let activeYoutubeSourceUrlPromise: Promise<string | null> | null = null;
  let generation = 0;

  const handleMediaPathChange = async (mediaPath: string): Promise<void> => {
    const currentGeneration = ++generation;
    const config = deps.getMediaCacheConfig();
    if (config.mode !== 'background') {
      activeYoutubeSourceUrl = null;
      activeYoutubeSourceUrlPromise = Promise.resolve(null);
      return;
    }

    if (!isYoutubeMediaPath(mediaPath)) {
      activeYoutubeSourceUrl = null;
    }

    const sourceUrlPromise = resolveYoutubeSourceUrl(deps, mediaPath);
    activeYoutubeSourceUrlPromise = sourceUrlPromise.then((sourceUrl) =>
      currentGeneration === generation ? sourceUrl : null,
    );

    const sourceUrl = await sourceUrlPromise;
    if (currentGeneration !== generation) {
      return;
    }

    activeYoutubeSourceUrl = sourceUrl;
    if (!sourceUrl) {
      return;
    }

    try {
      deps.startYoutubeMediaCache(sourceUrl, {
        mode: config.mode,
        maxHeight: config.maxHeight,
      });
    } catch (error) {
      deps.logWarn(
        `Failed to start YouTube media cache: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }
  };

  return {
    getActiveYoutubeSourceUrl: async (): Promise<string | null> =>
      activeYoutubeSourceUrlPromise ?? activeYoutubeSourceUrl,
    getActiveYoutubeSourceUrlSnapshot: (): string | null => activeYoutubeSourceUrl,
    handleMediaPathChange,
  };
}
