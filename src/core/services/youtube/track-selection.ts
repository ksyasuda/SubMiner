import { isEnglishYoutubeLang, isJapaneseYoutubeLang } from './labels';
import type { YoutubeTrackOption } from './track-probe';

function pickTrack(
  tracks: YoutubeTrackOption[],
  matcher: (value: string) => boolean,
  excludeId?: string,
): YoutubeTrackOption | null {
  const matching = tracks.filter((track) => matcher(track.language) && track.id !== excludeId);
  return matching[0] ?? null;
}

export function chooseDefaultYoutubeTrackIds(tracks: YoutubeTrackOption[]): {
  primaryTrackId: string | null;
  secondaryTrackId: string | null;
} {
  const primary =
    pickTrack(
      tracks.filter((track) => track.kind === 'manual'),
      isJapaneseYoutubeLang,
    ) ||
    pickTrack(
      tracks.filter((track) => track.kind === 'auto'),
      isJapaneseYoutubeLang,
    ) ||
    tracks.find((track) => track.kind === 'manual') ||
    tracks[0] ||
    null;

  const secondary =
    pickTrack(
      tracks.filter((track) => track.kind === 'manual'),
      isEnglishYoutubeLang,
      primary?.id ?? undefined,
    ) ||
    pickTrack(
      tracks.filter((track) => track.kind === 'auto'),
      isEnglishYoutubeLang,
      primary?.id ?? undefined,
    ) ||
    null;

  return {
    primaryTrackId: primary?.id ?? null,
    secondaryTrackId: secondary?.id ?? null,
  };
}

export function normalizeYoutubeTrackSelection(input: {
  primaryTrackId: string | null;
  secondaryTrackId: string | null;
}): {
  primaryTrackId: string | null;
  secondaryTrackId: string | null;
} {
  if (
    input.primaryTrackId &&
    input.secondaryTrackId &&
    input.primaryTrackId === input.secondaryTrackId
  ) {
    return {
      primaryTrackId: input.primaryTrackId,
      secondaryTrackId: null,
    };
  }
  return input;
}
