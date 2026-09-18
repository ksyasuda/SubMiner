import { createHash } from 'node:crypto';

const URL_SCHEME = /[a-z][a-z0-9+.-]*:\/\//i;
const QUERY_PAIR = /[?&][^=\s&#]+=/;
const CREDENTIAL_LABEL = /\b(?:api[_ -]?key|access[_ -]?token|x[_ -]?emby[_ -]?token)\b/i;

/** mpv can report the URL or its query-bearing basename before metadata arrives. */
export function sanitizeMediaTitle(value: string | null | undefined): string | null {
  const title = value?.trim();
  if (!title) return null;
  let decoded = title;
  try {
    decoded = decodeURIComponent(title);
  } catch {
    // Ordinary titles can contain a literal percent sign.
  }
  if (URL_SCHEME.test(decoded) || QUERY_PAIR.test(decoded) || CREDENTIAL_LABEL.test(decoded)) {
    return null;
  }
  return title;
}

export function resolveMediaLookupTarget(
  mediaPath: string | null,
  mediaTitle: string | null,
): string | null {
  return sanitizeMediaTitle(mediaPath) ?? sanitizeMediaTitle(mediaTitle);
}

/** Persistent identity, never a URL to use for authenticated media retrieval. */
export function toMediaIdentityPath(mediaPath: string): string {
  const value = mediaPath.trim();
  if (!URL_SCHEME.test(value)) return sanitizeMediaTitle(value) ?? '';
  try {
    const url = new URL(value);
    const jellyfinItem = url.pathname.match(
      /\/Videos\/([^/]+)\/(?:stream(?:\.[^/]*)?|master\.m3u8)\/?$/i,
    );
    if (jellyfinItem) return `jellyfin://${url.host}/item/${jellyfinItem[1]}`;
    url.username = '';
    url.password = '';
    const query = url.search;
    const queryFingerprint = /^#query-[a-f0-9]{64}$/.test(url.hash) ? url.hash : '';
    const youtubeId = /^(?:www\.|m\.)?youtube\.com$/i.test(url.hostname)
      ? url.searchParams.get('v')
      : null;
    url.search = '';
    url.hash = queryFingerprint;
    if (youtubeId) url.searchParams.set('v', youtubeId);
    else if (query) {
      // Different query-selected videos must not collapse into one stats entry.
      url.hash = `query-${createHash('sha256').update(query).digest('hex')}`;
    }
    return url.toString();
  } catch {
    return '';
  }
}
