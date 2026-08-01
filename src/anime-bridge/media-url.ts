/**
 * The bridge returns cover art and video URLs pointing at its own loopback
 * media proxy, but the origin it embeds is not always the port we actually
 * started it on. Rebase those onto the live bridge origin, and leave any
 * genuinely remote URL untouched.
 */

const PROXY_ROUTES = new Set(['image', 'video']);
const LOOPBACK_HOSTS = new Set(['127.0.0.1', 'localhost', '::1', '[::1]']);

function isLoopbackProxyUrl(candidate: URL): boolean {
  if (candidate.protocol !== 'http:' && candidate.protocol !== 'https:') return false;
  const host = candidate.hostname.toLowerCase();
  if (!LOOPBACK_HOSTS.has(host)) return false;
  const route = candidate.pathname.split('/').filter(Boolean)[0];
  return route !== undefined && PROXY_ROUTES.has(route);
}

/**
 * Rewrite a bridge media URL onto `bridgeBaseUrl`, preserving path and query.
 * Returns the input unchanged when it is not a loopback proxy URL, or when
 * either URL cannot be parsed.
 */
export function resolveBridgeMediaUrl(bridgeBaseUrl: string, mediaUrl: string): string {
  let media: URL;
  try {
    media = new URL(mediaUrl);
  } catch {
    return mediaUrl;
  }
  if (!isLoopbackProxyUrl(media)) return mediaUrl;

  const normalizedBase = bridgeBaseUrl.includes('://') ? bridgeBaseUrl : `http://${bridgeBaseUrl}`;
  let base: URL;
  try {
    base = new URL(normalizedBase);
  } catch {
    return mediaUrl;
  }
  if (base.protocol !== 'http:' && base.protocol !== 'https:') return mediaUrl;
  if (base.hostname.length === 0) return mediaUrl;

  const rebased = new URL(base.origin);
  rebased.pathname = media.pathname;
  rebased.search = media.search;
  rebased.hash = media.hash;
  return rebased.toString();
}

/**
 * Send a bridge-served HLS stream through the local strip proxy instead. Only
 * `.m3u8` URLs on the bridge origin qualify: direct files need no fixing, and
 * an external URL would not resolve through a proxy that forwards to the
 * bridge. Anything unparseable comes back unchanged.
 */
export function routeHlsThroughProxy(
  streamUrl: string,
  bridgeBaseUrl: string,
  proxyOrigin: string,
): string {
  let stream: URL;
  let bridge: URL;
  try {
    stream = new URL(streamUrl);
    bridge = new URL(bridgeBaseUrl);
  } catch {
    return streamUrl;
  }
  if (stream.origin !== bridge.origin) return streamUrl;
  if (!stream.pathname.endsWith('.m3u8')) return streamUrl;
  return `${proxyOrigin}${stream.pathname}${stream.search}`;
}

/** Aniyomi's SAnime status constants. */
export type AnimeStatus =
  | 'unknown'
  | 'ongoing'
  | 'completed'
  | 'publishing-finished'
  | 'cancelled'
  | 'on-hiatus';

export function parseAnimeStatus(status: number | undefined): AnimeStatus {
  switch (status) {
    case 1:
      return 'ongoing';
    case 2:
      return 'completed';
    case 4:
      return 'publishing-finished';
    case 5:
      return 'cancelled';
    case 6:
      return 'on-hiatus';
    default:
      return 'unknown';
  }
}
