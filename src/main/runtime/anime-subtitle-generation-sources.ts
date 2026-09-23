import type { ResolvedStream } from '../../anime-bridge/types';
import { sanitizeMediaHttpHeaders } from '../../core/services/mpv-http-headers';
import type { SubtitleGenerationAlternative } from '../../core/services/subtitle-generation-source';

function height(label: string): number | null {
  const match = /\b(\d{3,4})p\b/i.exec(label);
  return match ? Number(match[1]) : null;
}

/** These are candidates only. Generation verifies their audio against playback. */
export function animeSubtitleGenerationSources(
  selected: ResolvedStream,
  streams: readonly ResolvedStream[],
): SubtitleGenerationAlternative[] {
  const httpHeaders = { headers: sanitizeMediaHttpHeaders(selected.headers), userAgent: null };
  const audios: SubtitleGenerationAlternative[] = selected.audios
    .filter((track) => /^(?:ja|jpn|japanese)(?:[-_\s]|$)|日本語/i.test(track.lang))
    .map((track) => ({ kind: 'audio', url: track.url, label: 'Japanese audio', httpHeaders }));
  const selectedHeight = height(selected.quality);
  const videos: SubtitleGenerationAlternative[] = streams
    .filter((stream) => {
      const resolution = height(stream.quality);
      return (
        stream.url !== selected.url &&
        selectedHeight !== null &&
        resolution !== null &&
        resolution < selectedHeight &&
        !/\bdub(?:bed)?\b/i.test(stream.quality)
      );
    })
    .sort((a, b) => (height(a.quality) ?? Infinity) - (height(b.quality) ?? Infinity))
    .map((stream) => ({
      kind: 'video',
      url: stream.url,
      label: stream.quality,
      httpHeaders: { headers: sanitizeMediaHttpHeaders(stream.headers), userAgent: null },
    }));
  return [...audios, ...videos]
    .filter(
      (candidate, index, all) =>
        /^https?:\/\//i.test(candidate.url) &&
        all.findIndex((entry) => entry.url === candidate.url) === index,
    )
    .slice(0, 4);
}
