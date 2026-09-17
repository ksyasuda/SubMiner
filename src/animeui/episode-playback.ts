import type { AnimeBrowserPlaybackState } from '../types/anime-browser';

interface AnimeIdentity {
  sourceId: string;
  url: string;
}

export interface EpisodePlaybackCue {
  url: string;
  state: 'loading' | 'playing';
}

export function playingEpisodeForAnime(
  playback: AnimeBrowserPlaybackState | null,
  anime: AnimeIdentity | null,
): string | null {
  if (
    !playback ||
    !anime ||
    anime.sourceId !== playback.sourceId ||
    anime.url !== playback.animeUrl
  ) {
    return null;
  }
  return playback.episodeUrl;
}

export function nextPlaybackCue(
  playback: AnimeBrowserPlaybackState | null,
  anime: AnimeIdentity | null,
  currentCue: EpisodePlaybackCue | null,
): EpisodePlaybackCue | null {
  if (currentCue?.state === 'loading') return currentCue;
  const episodeUrl = playingEpisodeForAnime(playback, anime);
  return episodeUrl ? { url: episodeUrl, state: 'playing' } : null;
}
