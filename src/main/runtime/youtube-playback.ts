function trimToNull(value: string | null | undefined): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function matchesYoutubeHost(hostname: string, expectedHost: string): boolean {
  return hostname === expectedHost || hostname.endsWith(`.${expectedHost}`);
}

function extractYoutubeVideoId(mediaPath: string | null | undefined): string | null {
  const normalized = trimToNull(mediaPath);
  if (!normalized) {
    return null;
  }

  let parsed: URL;
  try {
    parsed = new URL(normalized);
  } catch {
    return null;
  }

  const host = parsed.hostname.toLowerCase();
  if (matchesYoutubeHost(host, 'youtu.be')) {
    return parsed.pathname.replace(/^\/+/, '').split('/')[0]?.trim() || null;
  }
  if (
    !matchesYoutubeHost(host, 'youtube.com') &&
    !matchesYoutubeHost(host, 'youtube-nocookie.com')
  ) {
    return null;
  }
  if (parsed.pathname === '/watch') {
    return parsed.searchParams.get('v')?.trim() || null;
  }
  const pathSegments = parsed.pathname.replace(/^\/+/, '').split('/');
  if (pathSegments[0] === 'shorts' || pathSegments[0] === 'embed') {
    return pathSegments[1]?.trim() || null;
  }
  return null;
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
    matchesYoutubeHost(host, 'youtu.be') ||
    matchesYoutubeHost(host, 'youtube.com') ||
    matchesYoutubeHost(host, 'youtube-nocookie.com')
  );
}

export function isSameYoutubeMediaPath(
  left: string | null | undefined,
  right: string | null | undefined,
): boolean {
  const leftId = extractYoutubeVideoId(left);
  const rightId = extractYoutubeVideoId(right);
  return Boolean(leftId && rightId && leftId === rightId);
}

export function shouldUseCachedYoutubeParsedCues(input: {
  videoPath: string | null | undefined;
  cachedMediaPath: string | null | undefined;
  cachedCueCount: number;
}): boolean {
  return input.cachedCueCount > 0 && isSameYoutubeMediaPath(input.videoPath, input.cachedMediaPath);
}

export function isYoutubePlaybackActive(
  currentMediaPath: string | null | undefined,
  currentVideoPath: string | null | undefined,
): boolean {
  return isYoutubeMediaPath(currentMediaPath) || isYoutubeMediaPath(currentVideoPath);
}
