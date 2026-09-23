import assert from 'node:assert/strict';
import { createServer, type RequestListener } from 'node:http';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  downloadSubtitleGenerationHls,
  downloadSubtitleGenerationHlsWindow,
} from './subtitle-generation-hls';

function playlist(count: number) {
  return (
    '#EXTM3U\n#EXT-X-TARGETDURATION:1\n' +
    Array.from({ length: count }, (_, i) => `#EXTINF:1,\nsegment${i}`).join('\n') +
    '\n#EXT-X-ENDLIST\n'
  );
}

async function fixture(
  handler: RequestListener,
  run: (url: string, directory: string) => Promise<void>,
) {
  const directory = await mkdtemp(path.join(tmpdir(), 'subtitle-hls-test-'));
  const server = createServer(handler);
  try {
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
    const address = server.address();
    assert.ok(address && typeof address === 'object');
    await run(`http://127.0.0.1:${address.port}`, directory);
  } finally {
    server.closeAllConnections();
    await new Promise<void>((resolve) => server.close(() => resolve()));
    await rm(directory, { recursive: true, force: true });
  }
}

const httpHeaders = { headers: { Referer: 'https://anime.example/' }, userAgent: 'SubMiner test' };

test('HLS workers preserve order, resolve redirected relative URLs and bound concurrency', async () => {
  let active = 0;
  let peak = 0;
  const completed: number[] = [];
  const progress: number[] = [];
  const bytes = (index: number) => {
    const value = Buffer.alloc(188 * 3, index);
    value[0] = value[188] = value[376] = 0x47;
    return value;
  };
  await fixture(
    async (req, res) => {
      if (
        req.headers.referer !== httpHeaders.headers.Referer ||
        req.headers['user-agent'] !== httpHeaders.userAgent
      ) {
        res.writeHead(403).end();
        return;
      }
      if (req.url === '/episode') {
        res.writeHead(302, { Location: '/media/list' }).end();
        return;
      }
      if (req.url === '/media/list') {
        res.end(playlist(10));
        return;
      }
      const match = /^\/media\/segment(\d+)$/.exec(req.url ?? '');
      if (!match) {
        res.writeHead(404).end();
        return;
      }
      const index = Number(match[1]);
      active += 1;
      peak = Math.max(active, peak);
      await new Promise((resolve) => setTimeout(resolve, index === 0 ? 80 : 10));
      completed.push(index);
      active -= 1;
      res.end(bytes(index));
    },
    async (url, directory) => {
      const result = await downloadSubtitleGenerationHls({
        mediaPath: `${url}/episode`,
        directory,
        httpHeaders,
        onProgress: (percent) => progress.push(percent),
      });
      assert.ok(result);
      assert.equal(peak, 4);
      assert.notEqual(completed[0], 0);
      assert.deepEqual(progress, [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]);
      const staged = await readFile(result, 'utf8');
      assert.equal(staged, playlist(10).replace(/segment(\d+)/g, 'segment-$1.ts'));
      for (let i = 0; i < 10; i += 1) {
        assert.deepEqual(
          await readFile(path.join(path.dirname(result), `segment-${i}.ts`)),
          bytes(i),
        );
      }
    },
  );
});

for (const mode of ['cancel', 'failure'] as const) {
  test(
    `HLS ${mode} aborts active downloads and drains workers before returning`,
    { timeout: 5000 },
    async () => {
      const controller = new AbortController();
      let requests = 0;
      let closed = 0;
      let triggerFailure = () => {};
      let allClosed = () => {};
      const closedPromise = new Promise<void>((resolve) => {
        allClosed = resolve;
      });
      await fixture(
        (req, res) => {
          if (req.url === '/episode') {
            res.end(playlist(12));
            return;
          }
          requests += 1;
          let settled = false;
          const settle = () => {
            if (settled) return;
            settled = true;
            if (++closed === 4) allClosed();
          };
          res.once('finish', settle);
          req.socket.once('close', settle);
          if (req.url === '/segment0') triggerFailure = () => res.writeHead(503).end();
          else res.write(Buffer.alloc(188));
          if (requests === 4) {
            if (mode === 'cancel') controller.abort();
            else triggerFailure();
          }
        },
        async (url, directory) => {
          await assert.rejects(
            downloadSubtitleGenerationHls({
              mediaPath: `${url}/episode`,
              directory,
              httpHeaders,
              signal: controller.signal,
            }),
            mode === 'cancel' ? /abort/i : /HTTP 503/,
          );
          await closedPromise;
          assert.equal(requests, 4);
          await assert.rejects(readFile(path.join(directory, 'hls', 'episode.m3u8')), /ENOENT/);
        },
      );
    },
  );
}

test('complex or live playlists fall back without downloading segments', async () => {
  const unsupported = [
    playlist(2).replace('#EXT-X-ENDLIST', ''),
    playlist(2).replace('#EXTINF:1,', '#EXT-X-KEY:METHOD=AES-128,URI="key"\n#EXTINF:1,'),
    playlist(2).replace('#EXTINF:1,', '#EXT-X-BYTERANGE:100@0\n#EXTINF:1,'),
    playlist(2).replace('#EXTINF:1,', '#EXT-X-MAP:URI="init.mp4"\n#EXTINF:1,'),
    playlist(2).replace('#EXTINF:1,', '#EXT-X-DISCONTINUITY\n#EXTINF:1,'),
    '#EXTM3U\n#EXT-X-STREAM-INF:BANDWIDTH=1000\nchild.m3u8\n',
    playlist(2).replace('segment0', 'file:///private/media.ts'),
    playlist(2).replace('#EXTINF:1,', '#EXTINF:NaN,'),
  ];
  let requests = 0;
  await fixture(
    (req, res) => {
      requests += 1;
      res.end(unsupported[Number(req.url?.slice(1))]);
    },
    async (url, directory) => {
      for (let i = 0; i < unsupported.length; i += 1) {
        assert.equal(
          await downloadSubtitleGenerationHls({ mediaPath: `${url}/${i}`, directory, httpHeaders }),
          null,
        );
      }
      assert.equal(requests, unsupported.length);
    },
  );
});

test('packed audio retains the original FFmpeg path', async () => {
  await fixture(
    (req, res) => {
      res.end(req.url === '/episode' ? playlist(1) : 'ID3 packed audio');
    },
    async (url, directory) => {
      assert.equal(
        await downloadSubtitleGenerationHls({
          mediaPath: `${url}/episode`,
          directory,
          httpHeaders,
        }),
        null,
      );
    },
  );
});

test('HLS samples download only segments covering the requested window', async () => {
  const requested: string[] = [];
  await fixture(
    (req, res) => {
      if (req.url === '/episode') {
        res.end(playlist(10));
        return;
      }
      requested.push(req.url ?? '');
      const bytes = Buffer.alloc(188 * 3);
      bytes[0] = bytes[188] = bytes[376] = 0x47;
      res.end(bytes);
    },
    async (url, directory) => {
      const result = await downloadSubtitleGenerationHlsWindow({
        mediaPath: `${url}/episode`,
        directory,
        httpHeaders,
        window: { startSeconds: 5.25, durationSeconds: 2 },
      });
      assert.ok(result);
      assert.deepEqual(requested.sort(), ['/segment5', '/segment6', '/segment7']);
      assert.equal(result.seekSeconds, 0.25);
      assert.match(await readFile(result.playlistPath, 'utf8'), /segment-5.ts/);
    },
  );
});
