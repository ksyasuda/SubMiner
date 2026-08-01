import type { BridgeVideo, OkHttpHeaders, ResolvedStream } from './types';

/**
 * Flatten OkHttp's alternating `[name, value, name, value]` array into a map.
 * A trailing name with no value is dropped rather than mapped to undefined.
 */
export function parseOkHttpHeaders(headers: OkHttpHeaders | undefined): Record<string, string> {
  const flat = headers?.['namesAndValues$okhttp'];
  if (!Array.isArray(flat)) return {};

  const parsed: Record<string, string> = {};
  for (let i = 0; i + 1 < flat.length; i += 2) {
    const name = flat[i];
    const value = flat[i + 1];
    if (typeof name === 'string' && typeof value === 'string') parsed[name] = value;
  }
  return parsed;
}

/**
 * Render headers as mpv's `--http-header-fields` string list. mpv splits
 * entries on commas, so commas inside a value must be escaped.
 */
export function toMpvHeaderFields(headers: Record<string, string>): string {
  return Object.entries(headers)
    .map(([name, value]) => `${name}: ${value.replace(/,/g, '\\,')}`)
    .join(',');
}

function normalizeTracks(
  tracks: Array<{ url?: string; lang?: string }> | undefined,
): Array<{ url: string; lang: string }> {
  if (!Array.isArray(tracks)) return [];
  return tracks
    .filter((track): track is { url: string; lang?: string } => typeof track.url === 'string')
    .map((track) => ({ url: track.url, lang: track.lang ?? '' }));
}

/**
 * Normalize a bridge video into a playable stream. Returns null when the
 * extension produced no `videoUrl`, which happens for entries it failed to
 * resolve.
 */
export function resolveStream(video: BridgeVideo): ResolvedStream | null {
  if (typeof video.videoUrl !== 'string' || video.videoUrl.length === 0) return null;

  return {
    url: video.videoUrl,
    quality: video.quality ?? '',
    headers: parseOkHttpHeaders(video.headers),
    subtitles: normalizeTracks(video.subtitleTracks),
    audios: normalizeTracks(video.audioTracks),
  };
}
