import assert from 'node:assert/strict';
import test from 'node:test';
import { createJellyfinSubtitleCacheIo } from './jellyfin-subtitle-cache-io';

test('jellyfin subtitle cache io downloads tracks to temp files and cleans cache dirs', async () => {
  const writes: Array<{ filePath: string; bytes: string }> = [];
  const removed: Array<{ dir: string; recursive: boolean; force: boolean }> = [];
  const cacheIo = createJellyfinSubtitleCacheIo({
    tmpDir: () => '/tmp',
    makeTempDir: async (prefix) => {
      assert.equal(prefix, '/tmp/subminer-jellyfin-subtitles-');
      return '/tmp/subminer-jellyfin-subtitles-abc';
    },
    writeFile: async (filePath, bytes) => {
      writes.push({ filePath, bytes: new TextDecoder().decode(bytes) });
    },
    removeDir: (dir, options) => {
      removed.push({ dir, ...options });
    },
    fetch: async () => ({
      ok: true,
      status: 200,
      arrayBuffer: async () => new TextEncoder().encode('subtitle body').buffer as ArrayBuffer,
    }),
  });

  const cached = await cacheIo.cacheSubtitleTrack({
    index: 7,
    deliveryUrl: 'https://example.test/Items/1/Subtitles/7/Stream.ass?api_key=secret',
  });
  cacheIo.cleanupCachedSubtitles([cached.cleanupDir]);

  assert.deepEqual(cached, {
    path: '/tmp/subminer-jellyfin-subtitles-abc/track-7.ass',
    cleanupDir: '/tmp/subminer-jellyfin-subtitles-abc',
  });
  assert.deepEqual(writes, [
    {
      filePath: '/tmp/subminer-jellyfin-subtitles-abc/track-7.ass',
      bytes: 'subtitle body',
    },
  ]);
  assert.deepEqual(removed, [
    { dir: '/tmp/subminer-jellyfin-subtitles-abc', recursive: true, force: true },
  ]);
});

test('jellyfin subtitle cache io removes temp dir when download fails', async () => {
  const removed: string[] = [];
  const cacheIo = createJellyfinSubtitleCacheIo({
    tmpDir: () => '/tmp',
    makeTempDir: async () => '/tmp/subminer-jellyfin-subtitles-failed',
    writeFile: async () => {},
    removeDir: (dir) => {
      removed.push(dir);
    },
    fetch: async () => ({
      ok: false,
      status: 500,
      arrayBuffer: async () => new ArrayBuffer(0),
    }),
  });

  await assert.rejects(
    () => cacheIo.cacheSubtitleTrack({ index: 1, deliveryUrl: 'https://example.test/sub.srt' }),
    /HTTP 500/,
  );
  assert.deepEqual(removed, ['/tmp/subminer-jellyfin-subtitles-failed']);
});

test('jellyfin subtitle cache io awaits async temp cleanup when download fails', async () => {
  let removed = false;
  const cacheIo = createJellyfinSubtitleCacheIo({
    tmpDir: () => '/tmp',
    makeTempDir: async () => '/tmp/subminer-jellyfin-subtitles-failed',
    writeFile: async () => {},
    removeDir: async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
      removed = true;
    },
    fetch: async () => ({
      ok: false,
      status: 500,
      arrayBuffer: async () => new ArrayBuffer(0),
    }),
  });

  await assert.rejects(
    () => cacheIo.cacheSubtitleTrack({ index: 1, deliveryUrl: 'https://example.test/sub.srt' }),
    /HTTP 500/,
  );
  assert.equal(removed, true);
});
