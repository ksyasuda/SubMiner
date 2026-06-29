import assert from 'node:assert/strict';
import test from 'node:test';

import { createYoutubeMediaCachePlaybackRuntime } from './youtube-media-cache-playback';

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((innerResolve) => {
    resolve = innerResolve;
  });
  return { promise, resolve };
}

test('youtube media cache starts when mpv reports a youtube path in background mode', async () => {
  const starts: Array<{ url: string; maxHeight: number }> = [];
  const runtime = createYoutubeMediaCachePlaybackRuntime({
    getMediaCacheConfig: () => ({ mode: 'background', maxHeight: 480 }),
    startYoutubeMediaCache: (url, options) => {
      starts.push({ url, maxHeight: options.maxHeight ?? 0 });
    },
    logWarn: () => {},
  });

  await runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abc123');

  assert.deepEqual(starts, [
    {
      url: 'https://www.youtube.com/watch?v=abc123',
      maxHeight: 480,
    },
  ]);
});

test('youtube media cache starts from original playlist url when mpv path is a resolved stream', async () => {
  const starts: string[] = [];
  const propertyRequests: string[] = [];
  const runtime = createYoutubeMediaCachePlaybackRuntime({
    getMediaCacheConfig: () => ({ mode: 'background', maxHeight: 720 }),
    requestMpvProperty: async (name) => {
      propertyRequests.push(name);
      if (name === 'playlist-playing-pos') {
        return 0;
      }
      if (name === 'playlist') {
        return [{ filename: 'https://www.youtube.com/watch?v=abc123' }];
      }
      return null;
    },
    startYoutubeMediaCache: (url) => {
      starts.push(url);
    },
    logWarn: () => {},
  });

  await runtime.handleMediaPathChange(
    'https://rr1---sn.example.googlevideo.com/videoplayback?expire=1777777777',
  );

  assert.deepEqual(propertyRequests, ['playlist-playing-pos', 'playlist']);
  assert.deepEqual(starts, ['https://www.youtube.com/watch?v=abc123']);
  assert.equal(await runtime.getActiveYoutubeSourceUrl(), 'https://www.youtube.com/watch?v=abc123');
});

test('youtube media cache can recover playlist source from current playlist marker', async () => {
  const starts: string[] = [];
  const runtime = createYoutubeMediaCachePlaybackRuntime({
    getMediaCacheConfig: () => ({ mode: 'background', maxHeight: 720 }),
    requestMpvProperty: async (name) => {
      if (name === 'playlist-playing-pos') {
        return -1;
      }
      if (name === 'playlist') {
        return [
          {
            filename: 'https://example.com/other.mp4',
          },
          {
            filename: 'https://youtu.be/abc123',
            current: true,
          },
        ];
      }
      return null;
    },
    startYoutubeMediaCache: (url) => {
      starts.push(url);
    },
    logWarn: () => {},
  });

  await runtime.handleMediaPathChange(
    'https://rr1---sn.example.googlevideo.com/videoplayback?expire=1777777777',
  );

  assert.deepEqual(starts, ['https://youtu.be/abc123']);
  assert.equal(await runtime.getActiveYoutubeSourceUrl(), 'https://youtu.be/abc123');
});

test('youtube media source getter awaits in-flight playlist recovery', async () => {
  const starts: string[] = [];
  const playlist = createDeferred<unknown>();
  const runtime = createYoutubeMediaCachePlaybackRuntime({
    getMediaCacheConfig: () => ({ mode: 'background', maxHeight: 720 }),
    requestMpvProperty: async (name) => {
      if (name === 'playlist-playing-pos') {
        return 0;
      }
      if (name === 'playlist') {
        return await playlist.promise;
      }
      return null;
    },
    startYoutubeMediaCache: (url) => {
      starts.push(url);
    },
    logWarn: () => {},
  });

  const pathChange = runtime.handleMediaPathChange(
    'https://rr1---sn.example.googlevideo.com/videoplayback?expire=1777777777',
  );
  const sourceUrl = runtime.getActiveYoutubeSourceUrl();

  playlist.resolve([{ filename: 'https://www.youtube.com/watch?v=abc123' }]);

  assert.equal(await sourceUrl, 'https://www.youtube.com/watch?v=abc123');
  await pathChange;
  assert.deepEqual(starts, ['https://www.youtube.com/watch?v=abc123']);
});

test('youtube media cache ignores non-youtube paths and direct mode', async () => {
  const starts: string[] = [];
  const propertyRequests: string[] = [];
  let mode: 'direct' | 'background' = 'background';
  const runtime = createYoutubeMediaCachePlaybackRuntime({
    getMediaCacheConfig: () => ({ mode, maxHeight: 720 }),
    requestMpvProperty: async (name) => {
      propertyRequests.push(name);
      return [
        {
          filename: 'https://www.youtube.com/watch?v=abc123',
          current: true,
        },
      ];
    },
    startYoutubeMediaCache: (url) => {
      starts.push(url);
    },
    logWarn: () => {},
  });

  await runtime.handleMediaPathChange('/tmp/video.mkv');
  mode = 'direct';
  await runtime.handleMediaPathChange('https://youtu.be/abc123');

  assert.deepEqual(propertyRequests, []);
  assert.deepEqual(starts, []);
});

test('youtube media cache logs synchronous start failures from path changes', async () => {
  const warnings: string[] = [];
  const runtime = createYoutubeMediaCachePlaybackRuntime({
    getMediaCacheConfig: () => ({ mode: 'background', maxHeight: 720 }),
    startYoutubeMediaCache: () => {
      throw new Error('yt-dlp missing');
    },
    logWarn: (message) => {
      warnings.push(message);
    },
  });

  await runtime.handleMediaPathChange('https://youtu.be/abc123');

  assert.deepEqual(warnings, ['Failed to start YouTube media cache: yt-dlp missing']);
});
