function trimToNull(value: string | null | undefined): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export function isYoutubeMediaPath(mediaPath: string | null | undefined): boolean {
  const normalized = trimToNull(mediaPath);
  if (!normalized) {
    return false;
  }

  let parsed: URL;
  try {
    parsed = new URL(normalized);
  } catch {
    return false;
  }

  const host = parsed.hostname.toLowerCase();
  return (
    host === 'youtu.be' ||
    host.endsWith('.youtu.be') ||
    host.endsWith('youtube.com') ||
    host.endsWith('youtube-nocookie.com')
  );
}

export function isYoutubePlaybackActive(
  currentMediaPath: string | null | undefined,
  currentVideoPath: string | null | undefined,
): boolean {
  return isYoutubeMediaPath(currentMediaPath) || isYoutubeMediaPath(currentVideoPath);
}
