export interface PendingYoutubeMediaUpdate {
  sourceUrl: string;
  noteId: number;
  startTime: number;
  endTime: number;
  label: string | number;
  audioFieldName?: string;
  imageFieldName?: string;
  miscInfoFieldName?: string;
  generateAudio: boolean;
  generateImage: boolean;
  volumeScale?: number;
}

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function getYoutubeVideoId(rawUrl: string): string | null {
  try {
    const parsed = new URL(rawUrl);
    const host = parsed.hostname.toLowerCase();
    if (host === 'youtu.be' || host.endsWith('.youtu.be')) {
      return trimToNonEmptyString(parsed.pathname.replace(/^\/+/, '').split('/')[0]);
    }
    if (host === 'youtube.com' || host.endsWith('.youtube.com')) {
      const watchId = trimToNonEmptyString(parsed.searchParams.get('v'));
      if (watchId) {
        return watchId;
      }
      const parts = parsed.pathname.split('/').filter(Boolean);
      if (parts[0] === 'shorts' || parts[0] === 'embed' || parts[0] === 'live') {
        return trimToNonEmptyString(parts[1]);
      }
    }
  } catch {
    return null;
  }
  return null;
}

export function youtubeMediaUrlsMatch(a: string, b: string): boolean {
  if (a.trim() === b.trim()) {
    return true;
  }

  const aVideoId = getYoutubeVideoId(a);
  const bVideoId = getYoutubeVideoId(b);
  return Boolean(aVideoId && bVideoId && aVideoId === bVideoId);
}
