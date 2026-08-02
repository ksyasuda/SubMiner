import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as http from 'http';
import * as os from 'os';
import * as path from 'path';
import { cleanupTemporaryFile, extractSubtitleTrackToFile } from './subsync-extract';
import type { ResolvedMpvHttpHeaders } from './mpv-http-headers';

const SUBTITLE_BODY = '1\n00:00:01,000 --> 00:00:02,000\nhello\n';

interface StubServer {
  url: (pathname: string) => string;
  requests: Array<{ url: string; headers: http.IncomingHttpHeaders }>;
  close: () => Promise<void>;
}

async function startSubtitleServer(): Promise<StubServer> {
  const requests: StubServer['requests'] = [];
  const server = http.createServer((req, res) => {
    requests.push({ url: req.url ?? '', headers: req.headers });
    if (req.url?.startsWith('/forbidden')) {
      res.writeHead(403);
      res.end('nope');
      return;
    }
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.end(SUBTITLE_BODY);
  });

  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  const port = typeof address === 'object' && address ? address.port : 0;

  return {
    url: (pathname) => `http://127.0.0.1:${port}${pathname}`,
    requests,
    close: () =>
      new Promise<void>((resolve) => {
        server.close(() => resolve());
      }),
  };
}

test('extractSubtitleTrackToFile downloads an external track served over http', async () => {
  const server = await startSubtitleServer();
  try {
    const result = await extractSubtitleTrackToFile({
      resolveFfmpegPath: () => {
        throw new Error('external tracks must not resolve ffmpeg');
      },
      videoPath: server.url('/stream.mp4'),
      track: {
        id: 2,
        type: 'sub',
        external: true,
        'external-filename': server.url('/subs/ja.srt'),
      },
      httpHeaders: null,
    });

    assert.equal(result.temporary, true);
    assert.equal(path.extname(result.path), '.srt');
    assert.equal(fs.readFileSync(result.path, 'utf8'), SUBTITLE_BODY);

    cleanupTemporaryFile(result);
    assert.equal(fs.existsSync(result.path), false);
  } finally {
    await server.close();
  }
});

test('extractSubtitleTrackToFile forwards mpv request headers to the subtitle host', async () => {
  const server = await startSubtitleServer();
  const httpHeaders: ResolvedMpvHttpHeaders = {
    headers: { Referer: 'https://example.test/watch' },
    userAgent: 'SubMinerTest/1.0',
  };

  try {
    const result = await extractSubtitleTrackToFile({
      resolveFfmpegPath: () => {
        throw new Error('external tracks must not resolve ffmpeg');
      },
      videoPath: server.url('/stream.mp4'),
      track: {
        id: 2,
        type: 'sub',
        external: true,
        'external-filename': server.url('/subs/ja.srt'),
      },
      httpHeaders,
    });
    cleanupTemporaryFile(result);

    const sent = server.requests.at(-1);
    assert.equal(sent?.headers.referer, 'https://example.test/watch');
    assert.equal(sent?.headers['user-agent'], 'SubMinerTest/1.0');
  } finally {
    await server.close();
  }
});

test('extractSubtitleTrackToFile reports the HTTP status when the download fails', async () => {
  const server = await startSubtitleServer();
  try {
    await assert.rejects(
      extractSubtitleTrackToFile({
        resolveFfmpegPath: () => {
          throw new Error('external tracks must not resolve ffmpeg');
        },
        videoPath: server.url('/stream.mp4'),
        track: {
          id: 2,
          type: 'sub',
          external: true,
          'external-filename': server.url('/forbidden/ja.srt'),
        },
        httpHeaders: null,
      }),
      /Failed to download subtitle track.*403/s,
    );
  } finally {
    await server.close();
  }
});

test('extractSubtitleTrackToFile still uses a local external track in place', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-extract-'));
  const localPath = path.join(dir, 'local.srt');
  fs.writeFileSync(localPath, SUBTITLE_BODY);

  try {
    const result = await extractSubtitleTrackToFile({
      resolveFfmpegPath: () => {
        throw new Error('external tracks must not resolve ffmpeg');
      },
      videoPath: path.join(dir, 'video.mkv'),
      track: { id: 2, type: 'sub', external: true, 'external-filename': localPath },
      httpHeaders: null,
    });

    assert.deepEqual(result, { path: localPath, temporary: false });
    // A borrowed file must survive cleanup.
    cleanupTemporaryFile(result);
    assert.equal(fs.existsSync(localPath), true);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('extractSubtitleTrackToFile rejects a missing local external track', async () => {
  await assert.rejects(
    extractSubtitleTrackToFile({
      resolveFfmpegPath: () => {
        throw new Error('external tracks must not resolve ffmpeg');
      },
      videoPath: '/tmp/video.mkv',
      track: { id: 2, type: 'sub', external: true, 'external-filename': '/tmp/does-not-exist.srt' },
      httpHeaders: null,
    }),
    /Subtitle file not found/,
  );
});

test('internal WebVTT extraction uses ffmpeg webvtt muxer with a vtt output file', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-extract-'));
  const ffmpegPath = path.join(dir, 'ffmpeg-stub');
  const argsPath = path.join(dir, 'args.txt');
  fs.writeFileSync(
    ffmpegPath,
    `#!/bin/sh\nprintf '%s\\n' "$@" > '${argsPath}'\nfor arg in "$@"; do output="$arg"; done\n: > "$output"\n`,
    { mode: 0o755 },
  );

  try {
    const result = await extractSubtitleTrackToFile({
      resolveFfmpegPath: () => ffmpegPath,
      videoPath: path.join(dir, 'video.mkv'),
      track: { id: 2, type: 'sub', codec: 'webvtt', 'ff-index': 3 },
      httpHeaders: null,
    });
    const args = fs.readFileSync(argsPath, 'utf8').trimEnd().split('\n');
    const formatIndex = args.indexOf('-f');

    assert.equal(args[formatIndex + 1], 'webvtt');
    assert.equal(path.extname(result.path), '.vtt');
    cleanupTemporaryFile(result);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('cleanupTemporaryFile preserves the retimed output sharing the temp directory', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-extract-'));
  const sourcePath = path.join(dir, 'remote_track.srt');
  const retimedPath = path.join(dir, 'remote_track_retimed.srt');
  fs.writeFileSync(sourcePath, SUBTITLE_BODY);
  fs.writeFileSync(retimedPath, SUBTITLE_BODY);

  try {
    cleanupTemporaryFile({ path: sourcePath, temporary: true }, retimedPath);

    assert.equal(fs.existsSync(sourcePath), false);
    assert.equal(fs.existsSync(retimedPath), true);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('cleanupTemporaryFile keeps the file when it is itself the preserved output', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-extract-'));
  const inPlacePath = path.join(dir, 'remote_track.srt');
  fs.writeFileSync(inPlacePath, SUBTITLE_BODY);

  try {
    cleanupTemporaryFile({ path: inPlacePath, temporary: true }, inPlacePath);

    assert.equal(fs.existsSync(inPlacePath), true);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
