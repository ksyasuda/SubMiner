import type { MiddlewareHandler } from 'hono';

function isLoopbackUrl(url: URL): boolean {
  return (
    url.protocol === 'http:' &&
    !url.username &&
    !url.password &&
    ['127.0.0.1', 'localhost', '[::1]'].includes(url.hostname)
  );
}

/** Protect the local API even when a browser can reach the loopback listener. */
export const enforceStatsRequestSafety: MiddlewareHandler = async (c, next) => {
  const url = new URL(c.req.url);
  if (!isLoopbackUrl(url)) return c.body(null, 403);

  const host = c.req.header('host');
  if (host !== undefined) {
    if (!/^(localhost|127\.0\.0\.1|\[::1\])(?::[0-9]+)?$/i.test(host)) {
      return c.body(null, 403);
    }
    // Node derives the request URL from Host; Bun provides them independently.
    try {
      if (new URL(`http://${host}`).origin !== url.origin) return c.body(null, 403);
    } catch {
      return c.body(null, 403);
    }
  }

  // Compare the serialized origin exactly. Opaque origins and malformed values
  // containing credentials, paths, or multiple origins must not gain trust.
  const origin = c.req.header('origin');
  if (origin !== undefined && origin !== url.origin) return c.body(null, 403);
  const site = c.req.header('sec-fetch-site');
  if (site === 'cross-site' || site === 'same-site') return c.body(null, 403);

  if (!['GET', 'HEAD', 'OPTIONS'].includes(c.req.method) && c.req.raw.body !== null) {
    const contentType = c.req.header('content-type')?.split(';', 1)[0]?.trim().toLowerCase();
    if (contentType !== 'application/json') return c.body(null, 415);
  }
  await next();
};
