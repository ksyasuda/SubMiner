import { fileURLToPath } from 'node:url';

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
    return track.type === 'sub' && track.id === sid && track.external === true;
  }) as Record<string, unknown> | undefined;

  const externalFilename =
    typeof activeTrack?.['external-filename'] === 'string'
      ? activeTrack['external-filename'].trim()
      : '';
  return externalFilename || null;
}

export function resolveSubtitleSourcePath(source: string): string {
  return source.startsWith('file://') ? fileURLToPath(new URL(source)) : source;
}
