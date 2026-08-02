import test from 'node:test';
import assert from 'node:assert/strict';
import * as path from 'path';
import {
  cacheSubtitleTracks,
  removeSubtitleCache,
  resolveSubtitleExtension,
  sniffSubtitleExtension,
  subtitleExtensionFromUrl,
  type SubtitleCacheIo,
} from './subtitle-cache';

interface FakeIo extends SubtitleCacheIo {
  written: Map<string, string>;
  removed: string[];
  requests: Array<{ url: string; headers: Record<string, string> }>;
}

function fakeIo(bodies: Record<string, string | { status: number }>): FakeIo {
  const written = new Map<string, string>();
  const removed: string[] = [];
  const requests: Array<{ url: string; headers: Record<string, string> }> = [];

  return {
    written,
    removed,
    requests,
    async fetch(url, init) {
      requests.push({ url, headers: init.headers });
      const body = bodies[url];
      if (body === undefined) throw new Error(`unexpected fetch: ${url}`);
      if (typeof body !== 'string') {
        return { ok: false, status: body.status, arrayBuffer: async () => new ArrayBuffer(0) };
      }
      const bytes = Buffer.from(body, 'utf8');
      return {
        ok: true,
        status: 200,
        arrayBuffer: async () =>
          bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
      };
    },
    async makeTempDir(prefix) {
      return `${prefix}test`;
    },
    async writeFile(filePath, bytes) {
      written.set(filePath, Buffer.from(bytes).toString('utf8'));
    },
    async removeDir(dir) {
      removed.push(dir);
    },
  };
}

const SRT = '1\n00:00:01,000 --> 00:00:02,000\nこんにちは\n';
const ASS = '[Script Info]\nScriptType: v4.00+\n\n[Events]\n';

test('content decides the extension before the url does', () => {
  assert.equal(sniffSubtitleExtension(ASS), 'ass');
  assert.equal(sniffSubtitleExtension(SRT), 'srt');
  assert.equal(sniffSubtitleExtension('WEBVTT\n\n00:01.000 --> 00:02.000\n'), 'vtt');
  assert.equal(sniffSubtitleExtension('nothing recognisable'), null);
  // An ASS body served from a .srt URL keeps the extension its parser needs.
  assert.equal(resolveSubtitleExtension('http://host/sub.srt', ASS), 'ass');
});

test('a bom or leading whitespace does not hide the format marker', () => {
  assert.equal(sniffSubtitleExtension(`﻿${ASS}`), 'ass');
  assert.equal(sniffSubtitleExtension(`\n\n${SRT}`), 'srt');
});

test('the url extension is the fallback, and only for formats we know', () => {
  assert.equal(subtitleExtensionFromUrl('http://host/a/b.ASS?x=1'), 'ass');
  assert.equal(subtitleExtensionFromUrl('http://host/video/token'), null);
  assert.equal(subtitleExtensionFromUrl('http://host/a.mp4'), null);
  // Nothing to go on: mpv can still probe past a wrong name.
  assert.equal(resolveSubtitleExtension('http://host/video/token', 'unknown'), 'srt');
});

test('tracks are downloaded to a temp dir and handed back as file paths', async () => {
  const io = fakeIo({ 'http://bridge/sub/ja': SRT, 'http://bridge/sub/en': ASS });
  const result = await cacheSubtitleTracks({
    tracks: [
      { url: 'http://bridge/sub/ja', lang: 'Japanese' },
      { url: 'http://bridge/sub/en', lang: 'English' },
    ],
    headers: { Referer: 'https://host/' },
    io,
  });

  assert.ok(result.dir);
  assert.deepEqual(
    result.tracks.map((track) => path.basename(track.url)),
    ['track-0.srt', 'track-1.ass'],
  );
  assert.ok(result.tracks.every((track) => track.local));
  assert.equal(io.written.get(result.tracks[0]!.url), SRT);
  // The stream's headers ride along; some hosts gate the subtitle URL too.
  assert.deepEqual(io.requests[0]!.headers, { Referer: 'https://host/' });
});

test('a failed download keeps its url so the episode still plays', async () => {
  const io = fakeIo({ 'http://bridge/sub/ja': SRT, 'http://bridge/sub/en': { status: 404 } });
  const logged: string[] = [];
  const result = await cacheSubtitleTracks({
    tracks: [
      { url: 'http://bridge/sub/ja', lang: 'Japanese' },
      { url: 'http://bridge/sub/en', lang: 'English' },
    ],
    io,
    log: (message) => logged.push(message),
  });

  assert.equal(result.tracks[0]!.local, true);
  assert.equal(result.tracks[1]!.local, false);
  assert.equal(result.tracks[1]!.url, 'http://bridge/sub/en');
  assert.ok(logged.some((message) => message.includes('404')));
  // One track survived, so the directory stays.
  assert.ok(result.dir);
  assert.deepEqual(io.removed, []);
});

test('an oversized streamed subtitle stops early and falls back to its remote url', async () => {
  const io = fakeIo({});
  const logged: string[] = [];
  let chunksRead = 0;
  let buffered = false;
  const chunk = new Uint8Array(20 * 1024 * 1024);
  const body = new ReadableStream<Uint8Array>(
    {
      pull(controller) {
        chunksRead += 1;
        controller.enqueue(chunk);
      },
    },
    { highWaterMark: 0 },
  );
  io.fetch = async () => ({
    ok: true,
    status: 200,
    body,
    async arrayBuffer() {
      buffered = true;
      throw new Error('stream should not be buffered');
    },
  });

  const result = await cacheSubtitleTracks({
    tracks: [{ url: 'http://bridge/sub/oversized', lang: 'Japanese' }],
    io,
    log: (message) => logged.push(message),
  });

  assert.equal(buffered, false);
  assert.equal(chunksRead, 2);
  assert.equal(result.dir, null);
  assert.deepEqual(result.tracks, [
    {
      url: 'http://bridge/sub/oversized',
      lang: 'Japanese',
      sourceUrl: 'http://bridge/sub/oversized',
      local: false,
    },
  ]);
  assert.ok(logged.some((message) => message.includes('response too large')));
});

test('a directory with nothing in it is removed and not reported', async () => {
  const io = fakeIo({ 'http://bridge/sub/ja': { status: 500 } });
  const result = await cacheSubtitleTracks({
    tracks: [{ url: 'http://bridge/sub/ja', lang: 'Japanese' }],
    io,
  });

  assert.equal(result.dir, null);
  assert.equal(result.tracks[0]!.local, false);
  assert.equal(io.removed.length, 1);
});

test('duplicate and empty urls are dropped before anything is fetched', async () => {
  const io = fakeIo({ 'http://bridge/sub/ja': SRT });
  const result = await cacheSubtitleTracks({
    tracks: [
      { url: 'http://bridge/sub/ja', lang: 'Japanese' },
      { url: 'http://bridge/sub/ja', lang: 'Japanese' },
      { url: '', lang: 'English' },
    ],
    io,
  });

  assert.equal(result.tracks.length, 1);
  assert.equal(io.requests.length, 1);
});

test('no tracks means no temp directory at all', async () => {
  const io = fakeIo({});
  const result = await cacheSubtitleTracks({ tracks: [], io });

  assert.deepEqual(result, { dir: null, tracks: [] });
  assert.equal(io.written.size, 0);
});

test('cleanup is best effort and never throws', async () => {
  const io = fakeIo({});
  io.removeDir = async () => {
    throw new Error('EBUSY');
  };
  await removeSubtitleCache('/tmp/subminer-anime-subtitles-x', io);
  // A null directory is the common case after a source with no subtitles.
  await removeSubtitleCache(null, io);
});
