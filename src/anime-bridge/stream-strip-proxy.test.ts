import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import type { AddressInfo } from 'node:net';
import {
  findTsSyncOffset,
  rewritePlaylistOrigins,
  startStreamStripProxy,
  TS_SEGMENT_ALIAS_SUFFIX,
  TS_PACKET_LENGTH,
} from './stream-strip-proxy';

/** A valid MPEG-TS payload: 0x47 sync byte every 188 bytes. */
function makeTsPackets(count: number): Buffer {
  const data = Buffer.alloc(count * TS_PACKET_LENGTH, 0x11);
  for (let i = 0; i < count; i++) data[i * TS_PACKET_LENGTH] = 0x47;
  return data;
}

/** The disguise seen in the wild: a real PNG header glued before the TS data. */
const PNG_HEADER = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
]);

test('findTsSyncOffset returns 0 for clean TS data', () => {
  assert.equal(findTsSyncOffset(makeTsPackets(6)), 0);
});

test('findTsSyncOffset finds TS data behind a PNG prefix', () => {
  const disguised = Buffer.concat([PNG_HEADER, makeTsPackets(6)]);
  assert.equal(findTsSyncOffset(disguised), PNG_HEADER.length);
});

test('findTsSyncOffset ignores a lone sync byte in the junk prefix', () => {
  const junk = Buffer.alloc(64, 0x00);
  junk[10] = 0x47; // decoy: no packet run follows it
  const disguised = Buffer.concat([junk, makeTsPackets(6)]);
  assert.equal(findTsSyncOffset(disguised), junk.length);
});

test('findTsSyncOffset returns null when no TS run exists', () => {
  assert.equal(findTsSyncOffset(Buffer.alloc(4096, 0x42)), null);
});

test('findTsSyncOffset returns null when the run starts past the scan limit', () => {
  const disguised = Buffer.concat([Buffer.alloc(600, 0x00), makeTsPackets(6)]);
  assert.equal(findTsSyncOffset(disguised, 500), null);
});

test('rewritePlaylistOrigins swaps absolute upstream URLs and keeps relative lines', () => {
  const body = [
    '#EXTM3U',
    '#EXTINF:6.006,',
    '/video/aaa.ts',
    '#EXTINF:4.463,',
    'http://127.0.0.1:41569/video/bbb.ts',
  ].join('\n');
  const rewritten = rewritePlaylistOrigins(body, 'http://127.0.0.1:41569', 'http://127.0.0.1:9999');
  assert.ok(rewritten.includes('http://127.0.0.1:9999/video/bbb.ts'));
  assert.ok(rewritten.includes('\n/video/aaa.ts\n'));
  assert.ok(!rewritten.includes('41569'));
});

test('rewritePlaylistOrigins gives proxied disguised segments an ffmpeg-safe TS suffix', () => {
  const body = [
    '#EXTM3U',
    '#EXTINF:6.006,',
    '/video/relative.image?token=one',
    '#EXTINF:4.463,',
    'http://127.0.0.1:41569/video/absolute.image',
    '#EXTINF:3,',
    'https://cdn.example/video/external.image',
    '#EXTINF:2,',
    '/video/rotating.jpg',
    '#EXTINF:2,',
    '/video/rotating.css',
    '#EXTINF:2,',
    '/video/rotating.html',
    '#EXTINF:2,',
    '/video/extensionless',
  ].join('\n');
  const rewritten = rewritePlaylistOrigins(body, 'http://127.0.0.1:41569', 'http://127.0.0.1:9999');

  assert.ok(rewritten.includes(`/video/relative.image${TS_SEGMENT_ALIAS_SUFFIX}?token=one`));
  assert.ok(
    rewritten.includes(`http://127.0.0.1:9999/video/absolute.image${TS_SEGMENT_ALIAS_SUFFIX}`),
  );
  assert.ok(rewritten.includes('https://cdn.example/video/external.image'));
  assert.ok(rewritten.includes(`/video/rotating.jpg${TS_SEGMENT_ALIAS_SUFFIX}`));
  assert.ok(rewritten.includes(`/video/rotating.css${TS_SEGMENT_ALIAS_SUFFIX}`));
  assert.ok(rewritten.includes(`/video/rotating.html${TS_SEGMENT_ALIAS_SUFFIX}`));
  assert.ok(rewritten.includes(`/video/extensionless${TS_SEGMENT_ALIAS_SUFFIX}`));
});

test('rewritePlaylistOrigins leaves recognized media extensions unaliased', () => {
  const body = [
    '#EXTM3U',
    '#EXTINF:6,',
    '/video/plain.ts',
    '#EXTINF:6,',
    '/video/fragmented.m4s',
    '#EXTINF:6,',
    '/subs/line.vtt',
    '/video/nested.m3u8',
  ].join('\n');
  const rewritten = rewritePlaylistOrigins(body, 'http://127.0.0.1:41569', 'http://127.0.0.1:9999');
  assert.ok(!rewritten.includes(TS_SEGMENT_ALIAS_SUFFIX));
});

/* ---------- proxy end-to-end ---------- */

type Route = { status: number; contentType: string; body: Buffer };

/** Either a static routing table or a handler, for upstreams that need one. */
async function withProxy(
  routes: Record<string, Route> | http.RequestListener,
  run: (proxyOrigin: string, upstreamOrigin: string) => Promise<void>,
): Promise<void> {
  const upstream = http.createServer((req, res) => {
    if (typeof routes === 'function') {
      routes(req, res);
      return;
    }
    const route = routes[req.url ?? ''];
    if (!route) {
      res.writeHead(404).end('missing');
      return;
    }
    res.writeHead(route.status, { 'content-type': route.contentType });
    res.end(route.body);
  });
  await new Promise<void>((resolve) => upstream.listen(0, '127.0.0.1', resolve));
  const upstreamOrigin = `http://127.0.0.1:${(upstream.address() as AddressInfo).port}`;

  const proxy = await startStreamStripProxy({
    upstreamOrigin: () => upstreamOrigin,
    retryDelayMs: 5,
  });
  try {
    await run(proxy.origin, upstreamOrigin);
  } finally {
    await proxy.close();
    await new Promise<void>((resolve) => upstream.close(() => resolve()));
  }
}

async function fetchBytes(url: string): Promise<{ status: number; body: Buffer }> {
  const response = await fetch(url);
  return { status: response.status, body: Buffer.from(await response.arrayBuffer()) };
}

test('proxy strips the PNG disguise off a segment', async () => {
  const ts = makeTsPackets(8);
  const disguised = Buffer.concat([PNG_HEADER, ts]);
  await withProxy(
    { '/video/seg.ts': { status: 200, contentType: 'video/mp2t', body: disguised } },
    async (origin) => {
      const { status, body } = await fetchBytes(`${origin}/video/seg.ts`);
      assert.equal(status, 200);
      assert.deepEqual(body, ts);
    },
  );
});

test('proxy passes a clean segment through unchanged', async () => {
  const ts = makeTsPackets(8);
  await withProxy(
    { '/video/seg.ts': { status: 200, contentType: 'video/mp2t', body: ts } },
    async (origin) => {
      const { body } = await fetchBytes(`${origin}/video/seg.ts`);
      assert.deepEqual(body, ts);
    },
  );
});

test('proxy leaves non-TS bodies alone', async () => {
  const vtt = Buffer.from('WEBVTT\n\n00:00.000 --> 00:01.000\nhello\n');
  await withProxy(
    { '/video/sub.vtt': { status: 200, contentType: 'text/vtt', body: vtt } },
    async (origin) => {
      const { body } = await fetchBytes(`${origin}/video/sub.vtt`);
      assert.deepEqual(body, vtt);
    },
  );
});

test('proxy rewrites absolute upstream playlist entries to its own origin', async () => {
  await withProxy(
    (req, res) => {
      if (req.url !== '/video/list.m3u8') {
        res.writeHead(404).end();
        return;
      }
      // The proxy rewrites the Host header to the upstream it dialled, so this
      // is that origin — the one the playlist must not leak to mpv.
      res.writeHead(200, { 'content-type': 'application/vnd.apple.mpegurl' });
      res.end(`#EXTM3U\n#EXTINF:6,\nhttp://${req.headers.host}/video/abs.ts\n`);
    },
    async (proxyOrigin, upstreamOrigin) => {
      const { body } = await fetchBytes(`${proxyOrigin}/video/list.m3u8`);
      const text = body.toString('utf8');
      assert.ok(text.includes(`${proxyOrigin}/video/abs.ts`));
      assert.ok(!text.includes(upstreamOrigin));
    },
  );
});

test('proxy aliases image segment URLs and removes the alias before fetching upstream', async () => {
  const ts = makeTsPackets(8);
  const disguised = Buffer.concat([PNG_HEADER, ts]);
  await withProxy(
    {
      '/video/list.m3u8': {
        status: 200,
        contentType: 'application/vnd.apple.mpegurl',
        body: Buffer.from('#EXTM3U\n#EXTINF:6,\n/video/seg.image\n'),
      },
      '/video/seg.image': { status: 200, contentType: 'image/png', body: disguised },
    },
    async (proxyOrigin) => {
      const playlist = await fetch(`${proxyOrigin}/video/list.m3u8`).then((response) =>
        response.text(),
      );
      const segmentPath = playlist.split('\n').find((line) => line.startsWith('/video/'));
      assert.equal(segmentPath, `/video/seg.image${TS_SEGMENT_ALIAS_SUFFIX}`);

      const { status, body } = await fetchBytes(`${proxyOrigin}${segmentPath}`);
      assert.equal(status, 200);
      assert.deepEqual(body, ts);
    },
  );
});

test('proxy strips even when the client asks for a byte range', async () => {
  // ffmpeg opens every HLS segment with `Range: bytes=0-`. The proxy drops the
  // header, so the upstream answers 200 with the whole body and the strip
  // applies; a 206 would have been forwarded untouched.
  const ts = makeTsPackets(8);
  const disguised = Buffer.concat([PNG_HEADER, ts]);
  await withProxy(
    (req, res) => {
      if (req.headers.range !== undefined) {
        res.writeHead(206, {
          'content-type': 'image/png',
          'content-range': `bytes 0-${disguised.length - 1}/${disguised.length}`,
        });
        res.end(disguised);
        return;
      }
      res.writeHead(200, { 'content-type': 'image/png' });
      res.end(disguised);
    },
    async (proxyOrigin) => {
      const response = await fetch(`${proxyOrigin}/video/seg.ts`, {
        headers: { Range: 'bytes=0-' },
      });
      const body = Buffer.from(await response.arrayBuffer());
      assert.equal(response.status, 200);
      assert.deepEqual(body, ts);
    },
  );
});

test('proxy forwards error statuses without touching the body', async () => {
  await withProxy(
    { '/video/gone.ts': { status: 404, contentType: 'text/plain', body: Buffer.from('nope') } },
    async (origin) => {
      const { status, body } = await fetchBytes(`${origin}/video/gone.ts`);
      assert.equal(status, 404);
      assert.equal(body.toString('utf8'), 'nope');
    },
  );
});

/**
 * Upstream that fails the first `failures` hits per path, then serves the
 * route. Mirrors the bridge right after an episode resolve: mpv's immediate
 * segment fetch errors, the same fetch a moment later works.
 */
async function withFlakyUpstream(
  failures: number,
  failStatus: number,
  route: Route,
  run: (proxyOrigin: string, hits: () => number) => Promise<void>,
): Promise<void> {
  let hits = 0;
  const upstream = http.createServer((_req, res) => {
    hits += 1;
    if (hits <= failures) {
      res.writeHead(failStatus, { 'content-type': 'text/plain' }).end('not ready');
      return;
    }
    res.writeHead(route.status, { 'content-type': route.contentType });
    res.end(route.body);
  });
  await new Promise<void>((resolve) => upstream.listen(0, '127.0.0.1', resolve));
  const upstreamOrigin = `http://127.0.0.1:${(upstream.address() as AddressInfo).port}`;
  const proxy = await startStreamStripProxy({
    upstreamOrigin: () => upstreamOrigin,
    retryDelayMs: 5,
  });
  try {
    await run(proxy.origin, () => hits);
  } finally {
    await proxy.close();
    await new Promise<void>((resolve) => upstream.close(() => resolve()));
  }
}

test('proxy retries a failed segment fetch once and serves the retry', async () => {
  const ts = makeTsPackets(8);
  await withFlakyUpstream(
    1,
    404,
    { status: 200, contentType: 'video/mp2t', body: ts },
    async (origin, hits) => {
      const { status, body } = await fetchBytes(`${origin}/video/seg.ts`);
      assert.equal(status, 200);
      assert.deepEqual(body, ts);
      assert.equal(hits(), 2);
    },
  );
});

test('proxy gives up after one retry and forwards the error', async () => {
  await withFlakyUpstream(
    Infinity,
    503,
    { status: 200, contentType: 'video/mp2t', body: makeTsPackets(8) },
    async (origin, hits) => {
      const { status } = await fetchBytes(`${origin}/video/seg.ts`);
      assert.equal(status, 503);
      assert.equal(hits(), 2);
    },
  );
});
