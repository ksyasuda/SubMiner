import { fileURLToPath } from 'node:url';

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

export function getActiveExternalSubtitleSource(
  trackListRaw: unknown,
  sidRaw: unknown,
): string | null {
  if (!Array.isArray(trackListRaw) || sidRaw == null) {
    return null;
  }

  const sid =
    typeof sidRaw === 'number' ? sidRaw : typeof sidRaw === 'string' ? Number(sidRaw) : null;
  if (sid == null || !Number.isFinite(sid)) {
    return null;
  }

  const activeTrack = trackListRaw.find((entry: unknown) => {
    if (!entry || typeof entry !== 'object') {
      return false;
    }
    const track = entry as Record<string, unknown>;
    return track.type === 'sub' && track.id === sid;
  }) as Record<string, unknown> | undefined;

  const externalFilename =
    typeof activeTrack?.['external-filename'] === 'string'
      ? activeTrack['external-filename'].trim()
      : '';
  return externalFilename || null;
}

export function resolveSubtitleSourcePath(source: string): string {
  if (!source.startsWith('file://')) {
    return source;
  }

  try {
    return fileURLToPath(new URL(source));
  } catch {
    return source;
  }
}

export function buildSubtitleSidebarSourceKey(
  videoPath: string,
  track: unknown,
  fallbackSourcePath?: string,
): string {
  const normalizedVideoPath = videoPath.trim();
  if (track && typeof track === 'object' && normalizedVideoPath) {
    const subtitleTrack = track as Record<string, unknown>;
    const trackId = parseTrackId(subtitleTrack.id);
    const ffIndex = parseTrackId(subtitleTrack['ff-index']);
    if (trackId !== null || ffIndex !== null) {
      return `internal:${normalizedVideoPath}:track:${trackId ?? 'unknown'}:ff:${ffIndex ?? 'unknown'}`;
    }
  }

  return fallbackSourcePath ?? normalizedVideoPath;
}
