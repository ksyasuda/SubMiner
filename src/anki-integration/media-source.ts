import { isRemoteMediaPath } from '../jimaku/utils';
import type { MpvClient } from '../types/runtime';

export type MediaGenerationKind = 'audio' | 'video';

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function extractUrlsFromMpvEdlSource(source: string): string[] {
  const matches = source.matchAll(/%\d+%(https?:\/\/.*?)(?=;!new_stream|;!global_tags|$)/gms);
  return [...matches]
    .map((match) => trimToNonEmptyString(match[1]))
    .filter((value): value is string => value !== null);
}

function classifyMediaUrl(url: string): MediaGenerationKind | null {
  try {
    const mime = new URL(url).searchParams.get('mime')?.toLowerCase() ?? '';
    if (mime.startsWith('audio/')) {
      return 'audio';
    }
    if (mime.startsWith('video/')) {
      return 'video';
    }
  } catch {
    // Ignore malformed URLs and fall back to stream order.
  }

  return null;
}

function resolvePreferredUrlFromMpvEdlSource(
  source: string,
  kind: MediaGenerationKind,
): string | null {
  const urls = extractUrlsFromMpvEdlSource(source);
  if (urls.length === 0) {
    return null;
  }

  const typedMatch = urls.find((url) => classifyMediaUrl(url) === kind);
  if (typedMatch) {
    return typedMatch;
  }

  // mpv EDL sources usually list audio streams first and video streams last, so
  // when classifyMediaUrl cannot identify a typed URL we fall back to stream order.
  return kind === 'audio' ? (urls[0] ?? null) : (urls[urls.length - 1] ?? null);
}

export async function resolveMediaGenerationInputPath(
  mpvClient: Pick<MpvClient, 'currentVideoPath' | 'requestProperty'> | null | undefined,
  kind: MediaGenerationKind = 'video',
): Promise<string | null> {
  const currentVideoPath = trimToNonEmptyString(mpvClient?.currentVideoPath);
  if (!currentVideoPath) {
    return null;
  }

  if (!isRemoteMediaPath(currentVideoPath) || !mpvClient?.requestProperty) {
    return currentVideoPath;
  }

  try {
    const streamOpenFilename = trimToNonEmptyString(
      await mpvClient.requestProperty('stream-open-filename'),
    );
    if (streamOpenFilename?.startsWith('edl://')) {
      return resolvePreferredUrlFromMpvEdlSource(streamOpenFilename, kind) ?? streamOpenFilename;
    }
    if (streamOpenFilename) {
      return streamOpenFilename;
    }
  } catch {
    // Fall back to the current path when mpv does not expose a resolved stream URL.
  }

  return currentVideoPath;
}
