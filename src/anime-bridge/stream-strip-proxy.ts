import http from 'node:http';
import type { AddressInfo } from 'node:net';

/**
 * Loopback proxy between mpv and the anime bridge that undoes segment
 * disguises. Some hosts prepend a real image header (a 1x1 PNG in the wild) to
 * every HLS segment so scrapers see "an image"; ffmpeg then probes the segment
 * as a picture and playback dies with "no audio or video data played". Aniyomi
 * strips this in its player; mpv needs the bytes fixed before it sees them.
 *
 * Only bridge-origin `.m3u8` streams are routed through here (see
 * anime-browser-runtime). Playlist bodies get their absolute upstream origins
 * rewritten so segment requests come back through the proxy; segment bodies are
 * scanned for the first genuine MPEG-TS packet run and any junk before it is
 * dropped. Anything that is not TS (fMP4, VTT, keys) passes through untouched.
 */

export const TS_PACKET_LENGTH = 188;
const TS_SYNC_BYTE = 0x47;
/**
 * Sync bytes that must repeat at exact packet spacing before an offset counts
 * as TS data. One or two matches happen by chance in binary data; five in a
 * row at 188-byte strides do not.
 */
const SYNC_RUN = 5;
/** A disguise prefix is small; give up scanning after this much. */
export const DEFAULT_SCAN_LIMIT_BYTES = 1024 * 1024;
/** Bytes needed to either find a run within the limit or rule one out. */
const DECISION_BYTES = DEFAULT_SCAN_LIMIT_BYTES + (SYNC_RUN - 1) * TS_PACKET_LENGTH + 1;

/**
 * First offset at which a confirmed MPEG-TS packet run starts, or null when
 * the data does not look like TS at all (within the scan limit).
 */
export function findTsSyncOffset(
  data: Buffer,
  scanLimit = DEFAULT_SCAN_LIMIT_BYTES,
): number | null {
  const lastConfirmable = data.length - (SYNC_RUN - 1) * TS_PACKET_LENGTH - 1;
  const end = Math.min(lastConfirmable, scanLimit);
  for (let offset = 0; offset <= end; offset++) {
    if (data[offset] !== TS_SYNC_BYTE) continue;
    let confirmed = true;
    for (let packet = 1; packet < SYNC_RUN; packet++) {
      if (data[offset + packet * TS_PACKET_LENGTH] !== TS_SYNC_BYTE) {
        confirmed = false;
        break;
      }
    }
    if (confirmed) return offset;
  }
  return null;
}

/**
 * Point absolute playlist entries at the proxy. Relative entries already
 * resolve against whatever origin served the playlist, so they need no help.
 */
export function rewritePlaylistOrigins(
  body: string,
  upstreamOrigin: string,
  proxyOrigin: string,
): string {
  return body.split(upstreamOrigin).join(proxyOrigin);
}

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

/** Response headers that must not be forwarded verbatim. */
const DROPPED_HEADERS = new Set([
  'connection',
  'keep-alive',
  'transfer-encoding',
  'content-length',
]);

function forwardableHeaders(headers: http.IncomingHttpHeaders): http.OutgoingHttpHeaders {
  const result: http.OutgoingHttpHeaders = {};
  for (const [name, value] of Object.entries(headers)) {
    if (value === undefined || DROPPED_HEADERS.has(name.toLowerCase())) continue;
    result[name] = value;
  }
  return result;
}

export function startStreamStripProxy(
  options: StreamStripProxyOptions,
): Promise<StreamStripProxyHandle> {
  const log = options.log ?? (() => {});
  const retryDelayMs = options.retryDelayMs ?? DEFAULT_RETRY_DELAY_MS;

  const server = http.createServer((req, res) => {
    if (req.method !== 'GET' && req.method !== 'HEAD') {
      res.writeHead(405).end();
      return;
    }

    let upstreamUrl: URL;
    try {
      upstreamUrl = new URL(req.url ?? '/', options.upstreamOrigin());
    } catch {
      res.writeHead(502).end();
      return;
    }

    const requestHeaders = forwardableHeaders(req.headers);
    delete requestHeaders.host;
    // Never forward Range: ffmpeg opens every segment with `bytes=0-`, the
    // bridge answers some of those 206, and a partial response cannot be
    // stripped (only full 200 bodies are). Byte ranges into a resource whose
    // bytes this proxy rewrites would be incoherent anyway.
    delete requestHeaders.range;
    res.on('error', () => {});

    requestUpstream(req, res, upstreamUrl, requestHeaders, 0);
  });

  /**
   * One delayed retry on a failed GET: right after an episode resolve, the
   * bridge (or the host behind it) can error on the very first segment
   * fetches and be fine a moment later — mpv treats a playlist full of failed
   * segments as a dead file and gives up for good.
   */
  function requestUpstream(
    req: http.IncomingMessage,
    res: http.ServerResponse,
    upstreamUrl: URL,
    requestHeaders: http.OutgoingHttpHeaders,
    attempt: number,
  ): void {
    const mayRetry = req.method === 'GET' && attempt === 0;
    const retry = (): void => {
      setTimeout(
        () => requestUpstream(req, res, upstreamUrl, requestHeaders, attempt + 1),
        retryDelayMs,
      );
    };

    const upstreamRequest = http.request(
      upstreamUrl,
      { method: req.method, headers: requestHeaders },
      (upstream) => {
        const status = upstream.statusCode ?? 502;
        if (status === 404 || status >= 500) {
          if (mayRetry) {
            log(`[stream-proxy] upstream ${status} for ${upstreamUrl.pathname}; retrying once`);
            upstream.resume();
            retry();
            return;
          }
          log(`[stream-proxy] upstream ${status} for ${upstreamUrl.pathname}`);
        }
        handleUpstreamResponse(req, res, upstream);
      },
    );
    upstreamRequest.on('error', (error) => {
      if (mayRetry) {
        log(`[stream-proxy] upstream request failed: ${String(error)}; retrying once`);
        retry();
        return;
      }
      log(`[stream-proxy] upstream request failed: ${String(error)}`);
      if (!res.headersSent) res.writeHead(502);
      res.end();
    });
    upstreamRequest.end();
  }

  function handleUpstreamResponse(
    req: http.IncomingMessage,
    res: http.ServerResponse,
    upstream: http.IncomingMessage,
  ): void {
    const status = upstream.statusCode ?? 502;
    const pathname = (req.url ?? '').split('?', 1)[0] ?? '';
    const contentType = String(upstream.headers['content-type'] ?? '');
    const isPlaylist = pathname.endsWith('.m3u8') || contentType.includes('mpegurl');

    upstream.on('error', () => res.destroy());

    // Only a full 200 body is safe to modify; everything else (errors, range
    // responses, HEAD) forwards untouched.
    if (status !== 200 || req.method === 'HEAD') {
      res.writeHead(status, forwardableHeaders(upstream.headers));
      upstream.pipe(res);
      return;
    }

    if (isPlaylist) {
      const chunks: Buffer[] = [];
      upstream.on('data', (chunk: Buffer) => chunks.push(chunk));
      upstream.on('end', () => {
        const body = rewritePlaylistOrigins(
          Buffer.concat(chunks).toString('utf8'),
          options.upstreamOrigin(),
          origin,
        );
        res.writeHead(status, {
          ...forwardableHeaders(upstream.headers),
          'content-length': Buffer.byteLength(body, 'utf8'),
        });
        res.end(body);
      });
      return;
    }

    stripSegment(res, upstream);
  }

  /**
   * Buffer just enough of the body to find (or rule out) a TS packet run,
   * drop everything before it, then stream the rest through untouched.
   */
  function stripSegment(res: http.ServerResponse, upstream: http.IncomingMessage): void {
    const chunks: Buffer[] = [];
    let buffered = 0;

    const respond = (data: Buffer, remainderFollows: boolean): void => {
      const offset = findTsSyncOffset(data) ?? 0;
      if (offset > 0) log(`[stream-proxy] stripped ${offset} disguise bytes off a segment`);
      const body = offset > 0 ? data.subarray(offset) : data;

      const headers = forwardableHeaders(upstream.headers);
      const upstreamLength = Number(upstream.headers['content-length']);
      if (remainderFollows) {
        if (Number.isFinite(upstreamLength)) headers['content-length'] = upstreamLength - offset;
      } else {
        headers['content-length'] = body.length;
      }

      res.writeHead(upstream.statusCode ?? 200, headers);
      res.write(body);
    };

    const onData = (chunk: Buffer): void => {
      chunks.push(chunk);
      buffered += chunk.length;
      if (buffered < DECISION_BYTES) return;
      upstream.off('data', onData);
      upstream.off('end', onEnd);
      respond(Buffer.concat(chunks), true);
      upstream.pipe(res);
    };
    const onEnd = (): void => {
      respond(Buffer.concat(chunks), false);
      res.end();
    };
    upstream.on('data', onData);
    upstream.on('end', onEnd);
  }

  let origin = '';

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
