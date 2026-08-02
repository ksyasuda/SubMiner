import type { AnimeStreamMetadata } from '../../anime-bridge/episode-metadata';
import type { AnilistMediaGuess } from '../../core/services/anilist/anilist-updater';
import type { JimakuMediaInfo } from '../../types';

/**
 * Holds what the anime browser resolved for the episode currently streaming.
 *
 * Consumers otherwise have to re-derive the series and episode from the mpv
 * title, and some of them only ever see the stream URL, which carries no title
 * at all. Keeping the resolved fields lets each of them ask instead of guess.
 */
export interface StreamPlaybackMetadataStore {
  set: (metadata: AnimeStreamMetadata) => void;
  clear: () => void;
  /**
   * The metadata for `mediaPath`, or null when the player has moved on to
   * something else. Matching on the path is what makes this self-expiring —
   * there is no teardown hook to miss.
   */
  match: (mediaPath: string | null) => AnimeStreamMetadata | null;
}

export function createStreamPlaybackMetadataStore(): StreamPlaybackMetadataStore {
  let current: AnimeStreamMetadata | null = null;

  return {
    set(metadata: AnimeStreamMetadata): void {
      current = metadata;
    },
    clear(): void {
      current = null;
    },
    match(mediaPath: string | null): AnimeStreamMetadata | null {
      if (!current) return null;
      const trimmed = typeof mediaPath === 'string' ? mediaPath.trim() : '';
      if (!trimmed) return null;
      // The stats path is accepted too: stats rewrites the volatile stream URL
      // to it, so callers reading from there still resolve.
      return trimmed === current.mediaPath || trimmed === current.statsPath ? current : null;
    },
  };
}

/**
 * Match metadata for the path a caller requested, falling back to the active
 * player path only when the caller has no path of its own.
 */
export function matchRequestedStreamPlaybackMetadata(
  store: StreamPlaybackMetadataStore,
  requestedMediaPath: string | null,
  currentMediaPath: string | null,
): AnimeStreamMetadata | null {
  return store.match(requestedMediaPath ?? currentMediaPath);
}

/** AniList counts whole episodes, so a special numbered 6.5 cannot drive it. */
function wholeEpisode(episode: number | null): number | null {
  return typeof episode === 'number' && Number.isInteger(episode) && episode > 0 ? episode : null;
}

/**
 * The subtitle modals' prefill. `high` confidence is honest here: these fields
 * came from the source's own listing rather than from a filename guess, so the
 * modals may search on them without waiting for the user to confirm.
 */
export function toJimakuMediaInfo(metadata: AnimeStreamMetadata): JimakuMediaInfo {
  const episode = wholeEpisode(metadata.episodeNumber);
  return {
    title: metadata.seriesTitle,
    season: metadata.seasonNumber,
    episode,
    confidence: episode !== null ? 'high' : 'low',
    filename: metadata.displayTitle,
    rawTitle: metadata.displayTitle,
  };
}

/**
 * The AniList guess, or null when there is no episode to report — in which case
 * the caller should fall back to parsing the title as usual.
 */
export function toAnilistMediaGuess(metadata: AnimeStreamMetadata): AnilistMediaGuess | null {
  const episode = wholeEpisode(metadata.episodeNumber);
  if (!metadata.seriesTitle || episode === null) return null;
  return {
    title: metadata.seriesTitle,
    season: metadata.seasonNumber,
    episode,
    source: 'stream',
  };
}
