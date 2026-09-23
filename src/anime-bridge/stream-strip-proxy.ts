import http from 'node:http';
import type { AddressInfo } from 'node:net';
import { handleUpstreamResponse, TS_SEGMENT_ALIAS_SUFFIX } from './stream-strip-response';
import { forwardableRequestHeaders, requestUpstream } from './stream-strip-transport';

export {
  DEFAULT_SCAN_LIMIT_BYTES,
  findTsSyncOffset,
  rewritePlaylistOrigins,
  TS_PACKET_LENGTH,
  TS_SEGMENT_ALIAS_SUFFIX,
} from './stream-strip-response';

/**
 * Loopback proxy between mpv and the anime bridge that removes disguised HLS
 * segment prefixes before ffmpeg sees them.
 */

export interface StreamStripProxyOptions {
  /** Read per request so a bridge restart on a new port keeps working. */
  upstreamOrigin: () => string;
  log?: (message: string) => void;
  /** Pause before the single retry of a failed upstream GET. */
  retryDelayMs?: number;
}

const DEFAULT_RETRY_DELAY_MS = 400;

export interface StreamStripProxyHandle {
  origin: string;
  port: number;
  close: () => Promise<void>;
}

function resolveUpstreamUrl(requestTarget: string, configuredOrigin: string): URL {
  if (
    !requestTarget.startsWith('/') ||
    requestTarget.startsWith('//') ||
    requestTarget.includes('#')
  ) {
    throw new Error('Stream proxy requests must use origin-form targets.');
  }

  const upstream = new URL(configuredOrigin);
  if (upstream.protocol !== 'http:') {
    throw new Error('Stream proxy upstream must use HTTP.');
  }
  const resolved = new URL(requestTarget, upstream.origin);
  if (resolved.protocol !== 'http:' || resolved.origin !== upstream.origin) {
    throw new Error('Stream proxy request escaped the configured upstream origin.');
  }
  if (resolved.pathname.endsWith(TS_SEGMENT_ALIAS_SUFFIX)) {
    resolved.pathname = resolved.pathname.slice(0, -TS_SEGMENT_ALIAS_SUFFIX.length);
  }
  return resolved;
}

export function startStreamStripProxy(
  options: StreamStripProxyOptions,
): Promise<StreamStripProxyHandle> {
  const log = options.log ?? (() => {});
  const retryDelayMs = options.retryDelayMs ?? DEFAULT_RETRY_DELAY_MS;
  let origin = '';

  const server = http.createServer((req, res) => {
    if (req.method !== 'GET' && req.method !== 'HEAD') {
      res.writeHead(405).end();
      return;
    }

    let upstreamUrl: URL;
    try {
      upstreamUrl = resolveUpstreamUrl(req.url ?? '/', options.upstreamOrigin());
    } catch {
      res.writeHead(502).end();
      return;
    }

    const requestHeaders = forwardableRequestHeaders(req.headers);
    delete requestHeaders.host;
    // Rewritten bodies cannot honor byte ranges into the original representation.
    delete requestHeaders.range;
    res.on('error', () => {});

    requestUpstream({
      req,
      res,
      upstreamUrl,
      requestHeaders,
      retryDelayMs,
      log,
      handleResponse: (upstream) =>
        handleUpstreamResponse({
          req,
          res,
          upstream,
          upstreamOrigin: options.upstreamOrigin,
          proxyOrigin: () => origin,
          log,
        }),
    });
  });

  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address() as AddressInfo;
      origin = `http://127.0.0.1:${port}`;
      resolve({
        origin,
        port,
        close: () =>
          new Promise<void>((resolveClose) => {
            server.closeAllConnections?.();
            server.close(() => resolveClose());
          }),
      });
    });
  });
}
