const COVER_RETRY_PARAM = 'coverRetry';

export function appendCoverRetryToken(src: string, retryToken = 0): string {
  if (!Number.isFinite(retryToken) || retryToken <= 0) return src;

  const normalizedToken = String(Math.trunc(retryToken));
  try {
    const url = new URL(src, 'http://subminer.local');
    url.searchParams.set(COVER_RETRY_PARAM, normalizedToken);
    if (src.startsWith('/')) {
      return `${url.pathname}${url.search}${url.hash}`;
    }
    return url.toString();
  } catch {
    const separator = src.includes('?') ? '&' : '?';
    return `${src}${separator}${COVER_RETRY_PARAM}=${encodeURIComponent(normalizedToken)}`;
  }
}

export function getCoverRetryDelayMs(retryToken: number): number {
  return Math.min(30_000, 2_000 * 2 ** Math.min(Math.max(retryToken, 0), 4));
}
