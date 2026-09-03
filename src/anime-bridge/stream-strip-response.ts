import type http from 'node:http';

export const TS_PACKET_LENGTH = 188;
const TS_SYNC_BYTE = 0x47;
/** Five sync bytes at exact packet spacing make an accidental match implausible. */
const SYNC_RUN = 5;

export const TS_SEGMENT_ALIAS_SUFFIX = '.subminer.ts';
const FFMPEG_SAFE_SEGMENT_EXTENSIONS = new Set([
  '3gp',
  'aac',
  'avi',
  'ac3',
  'eac3',
  'flac',
  'mkv',
  'm3u8',
  'm4a',
  'm4s',
  'm4v',
  'mpg',
  'mov',
  'mp2',
  'mp3',
  'mp4',
  'mpeg',
  'mpegts',
  'ogg',
  'ogv',
  'oga',
  'ts',
  'vob',
  'vtt',
  'wav',
  'webvtt',
  'cmfv',
  'cmfa',
  'ec3',
  'fmp4',
]);

function needsTsSegmentAlias(pathname: string): boolean {
  const name = pathname.slice(pathname.lastIndexOf('/') + 1).toLowerCase();
  const dot = name.lastIndexOf('.');
  if (dot === -1) return true;
  return !FFMPEG_SAFE_SEGMENT_EXTENSIONS.has(name.slice(dot + 1));
}

export const DEFAULT_SCAN_LIMIT_BYTES = 1024 * 1024;
const DECISION_BYTES = DEFAULT_SCAN_LIMIT_BYTES + (SYNC_RUN - 1) * TS_PACKET_LENGTH + 1;

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

export function rewritePlaylistOrigins(
  body: string,
  upstreamOrigin: string,
  proxyOrigin: string,
): string {
  const rebased = body.split(upstreamOrigin).join(proxyOrigin);
  return rebased
    .split(/(\r?\n)/)
    .map((line) => {
      const uri = line.trim();
      if (!uri || uri.startsWith('#')) return line;

      let resolved: URL;
      try {
        resolved = new URL(uri, proxyOrigin);
      } catch {
        return line;
      }
      if (resolved.origin !== proxyOrigin || !needsTsSegmentAlias(resolved.pathname)) {
        return line;
      }

      const queryIndex = uri.search(/[?#]/);
      const aliasIndex = queryIndex === -1 ? uri.length : queryIndex;
      const leadingWhitespace = line.slice(0, line.indexOf(uri));
      const trailingWhitespace = line.slice(leadingWhitespace.length + uri.length);
      return `${leadingWhitespace}${uri.slice(0, aliasIndex)}${TS_SEGMENT_ALIAS_SUFFIX}${uri.slice(aliasIndex)}${trailingWhitespace}`;
    })
    .join('');
}

const DROPPED_RESPONSE_HEADERS = new Set([
  'connection',
  'keep-alive',
  'transfer-encoding',
  'content-length',
]);

function forwardableResponseHeaders(headers: http.IncomingHttpHeaders): http.OutgoingHttpHeaders {
  const result: http.OutgoingHttpHeaders = {};
  for (const [name, value] of Object.entries(headers)) {
    if (value === undefined || DROPPED_RESPONSE_HEADERS.has(name.toLowerCase())) continue;
    result[name] = value;
  }
  return result;
}

export function handleUpstreamResponse(options: {
  req: http.IncomingMessage;
  res: http.ServerResponse;
  upstream: http.IncomingMessage;
  upstreamOrigin: () => string;
  proxyOrigin: () => string;
  log: (message: string) => void;
}): void {
  const { req, res, upstream, log } = options;
  const status = upstream.statusCode ?? 502;
  const pathname = (req.url ?? '').split('?', 1)[0] ?? '';
  const contentType = String(upstream.headers['content-type'] ?? '');
  const isPlaylist = pathname.endsWith('.m3u8') || contentType.includes('mpegurl');

  upstream.on('error', () => res.destroy());

  if (status !== 200 || req.method === 'HEAD') {
    res.writeHead(status, forwardableResponseHeaders(upstream.headers));
    upstream.pipe(res);
    return;
  }

  if (isPlaylist) {
    const chunks: Buffer[] = [];
    let buffered = 0;
    upstream.on('data', (chunk: Buffer) => {
      buffered += chunk.length;
      if (buffered > DECISION_BYTES) {
        log(`[stream-proxy] playlist body over ${DECISION_BYTES} bytes; dropping`);
        upstream.destroy();
        res.destroy();
        return;
      }
      chunks.push(chunk);
    });
    upstream.on('end', () => {
      const body = rewritePlaylistOrigins(
        Buffer.concat(chunks).toString('utf8'),
        options.upstreamOrigin(),
        options.proxyOrigin(),
      );
      res.writeHead(status, {
        ...forwardableResponseHeaders(upstream.headers),
        'content-length': Buffer.byteLength(body, 'utf8'),
      });
      res.end(body);
    });
    return;
  }

  stripSegment(res, upstream, log);
}

function stripSegment(
  res: http.ServerResponse,
  upstream: http.IncomingMessage,
  log: (message: string) => void,
): void {
  const chunks: Buffer[] = [];
  let buffered = 0;

  const respond = (data: Buffer, remainderFollows: boolean): void => {
    const offset = findTsSyncOffset(data) ?? 0;
    if (offset > 0) log(`[stream-proxy] stripped ${offset} disguise bytes off a segment`);
    const body = offset > 0 ? data.subarray(offset) : data;

    const headers = forwardableResponseHeaders(upstream.headers);
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
