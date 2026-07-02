import assert from 'node:assert/strict';
import test from 'node:test';

import * as mediaSource from './media-source';
import { toMpvEdlValue } from './mpv-edl-test-utils';

const { resolveMediaGenerationInputPath } = mediaSource;

type StructuredMediaInput = {
  path: string;
  source: string;
  singleResolvedStream: boolean;
  inputOptions?: {
    reconnect?: boolean;
    userAgent?: string;
    headers?: Record<string, string>;
  };
};

type StructuredMediaResolver = (
  mpvClient: Parameters<typeof resolveMediaGenerationInputPath>[0],
  kind?: Parameters<typeof resolveMediaGenerationInputPath>[1],
  options?: {
    getCachedMediaPath?: (
      currentVideoPath: string,
      kind: Parameters<typeof resolveMediaGenerationInputPath>[1],
    ) => Promise<string | null>;
    remoteCacheMode?: 'optional' | 'required';
    logDebug?: (message: string) => void;
  },
) => Promise<StructuredMediaInput | null>;

test('resolveMediaGenerationInputPath keeps local file paths', async () => {
  const result = await resolveMediaGenerationInputPath({
    currentVideoPath: '/tmp/video.mkv',
  });

  assert.equal(result, '/tmp/video.mkv');
});

test('resolveMediaGenerationInputPath prefers stream-open-filename for remote media', async () => {
  const requests: string[] = [];

  const result = await resolveMediaGenerationInputPath({
    currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
    requestProperty: async (name: string) => {
      requests.push(name);
      return 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123';
    },
  });

  assert.equal(result, 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123');
  assert.deepEqual(requests, ['stream-open-filename']);
});

test('resolveMediaGenerationInputPath unwraps mpv edl source for audio and video', async () => {
  const audioUrl = 'https://audio.example/videoplayback?mime=audio%2Fwebm';
  const videoUrl = 'https://video.example/videoplayback?mime=video%2Fmp4';
  const edlSource = [
    `edl://!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(audioUrl)}`,
    `!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(videoUrl)}`,
    '!global_tags,title=test',
  ].join(';');

  const audioResult = await resolveMediaGenerationInputPath(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      requestProperty: async () => edlSource,
    },
    'audio',
  );
  const videoResult = await resolveMediaGenerationInputPath(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      requestProperty: async () => edlSource,
    },
    'video',
  );

  assert.equal(audioResult, audioUrl);
  assert.equal(videoResult, videoUrl);
});

test('resolveMediaGenerationInputPath strips mpv edl segment options from unwrapped streams', async () => {
  const audioUrl = 'https://audio.example/videoplayback?mime=audio%2Fwebm';
  const signedVideoUrl =
    'https://rr1---sn.example.googlevideo.com/videoplayback?mime=video%2Fmp4&mn=sn-a,sn-b&lsig=abc%3D';
  const edlSource = [
    `edl://!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(audioUrl)}`,
    `!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(signedVideoUrl)},title=clip,length=73,timestamps=chapters`,
    '!global_tags,title=test',
  ].join(';');

  const result = await resolveMediaGenerationInputPath(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      requestProperty: async () => edlSource,
    },
    'video',
  );

  assert.equal(result, signedVideoUrl);
});

test('resolveMediaGenerationInputPath ignores length-guarded URLs in mpv edl headers', async () => {
  const initUrl = 'https://init.example/init.mp4';
  const audioUrl = 'https://audio.example/stream';
  const videoUrl = 'https://video.example/stream';
  const edlSource = [
    `edl://!mp4_dash,init=${toMpvEdlValue(initUrl)}`,
    '!new_stream',
    toMpvEdlValue(audioUrl),
    '!new_stream',
    toMpvEdlValue(videoUrl),
  ].join(';');

  const audioResult = await resolveMediaGenerationInputPath(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      requestProperty: async () => edlSource,
    },
    'audio',
  );

  assert.equal(audioResult, audioUrl);
});

test('resolveMediaGenerationInputPath falls back to currentVideoPath when stream-open-filename fails', async () => {
  const result = await resolveMediaGenerationInputPath({
    currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
    requestProperty: async () => {
      throw new Error('property unavailable');
    },
  });

  assert.equal(result, 'https://www.youtube.com/watch?v=abc123');
});

test('resolveMediaGenerationInput returns single-stream metadata for mpv EDL URLs', async () => {
  const resolver = (
    mediaSource as typeof mediaSource & {
      resolveMediaGenerationInput?: StructuredMediaResolver;
    }
  ).resolveMediaGenerationInput;
  assert.equal(typeof resolver, 'function');

  const audioUrl = 'https://audio.example/videoplayback?mime=audio%2Fwebm';
  const videoUrl = 'https://video.example/videoplayback?mime=video%2Fmp4';
  const edlSource = [
    `edl://!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(audioUrl)}`,
    `!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(videoUrl)}`,
  ].join(';');

  const result = await resolver!(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      requestProperty: async (name: string) => {
        if (name === 'stream-open-filename') return edlSource;
        if (name === 'user-agent') return 'Mozilla/5.0';
        if (name === 'http-header-fields') {
          return ['Cookie: SID=secret', 'Referer: https://www.youtube.com/', 'X-Test: ok'];
        }
        return null;
      },
    },
    'audio',
  );

  assert.equal(result?.path, audioUrl);
  assert.equal(result?.singleResolvedStream, true);
  assert.equal(result?.inputOptions?.reconnect, true);
  assert.equal(result?.inputOptions?.userAgent, 'Mozilla/5.0');
  assert.deepEqual(result?.inputOptions?.headers, {
    Referer: 'https://www.youtube.com/',
    'X-Test': 'ok',
  });
});

test('resolveMediaGenerationInput reads file-local mpv request options', async () => {
  const resolver = (
    mediaSource as typeof mediaSource & {
      resolveMediaGenerationInput?: StructuredMediaResolver;
    }
  ).resolveMediaGenerationInput;
  assert.equal(typeof resolver, 'function');

  const result = await resolver!(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      requestProperty: async (name: string) => {
        if (name === 'stream-open-filename') {
          return 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123';
        }
        if (name === 'file-local-options/user-agent') return 'SubMiner Test Agent';
        if (name === 'options/http-header-fields') return ['X-Shared: ok'];
        if (name === 'file-local-options/http-header-fields') {
          return ['Cookie: SID=secret', 'Referer: https://m.youtube.com/', 'X-Local: yes'];
        }
        return null;
      },
    },
    'video',
  );

  assert.equal(result?.path, 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123');
  assert.equal(result?.singleResolvedStream, true);
  assert.equal(result?.inputOptions?.userAgent, 'SubMiner Test Agent');
  assert.deepEqual(result?.inputOptions?.headers, {
    'X-Shared': 'ok',
    Referer: 'https://m.youtube.com/',
    'X-Local': 'yes',
    Origin: 'https://www.youtube.com',
  });
});

test('resolveMediaGenerationInput prefers a ready cached media file for YouTube extraction', async () => {
  const resolver = (
    mediaSource as typeof mediaSource & {
      resolveMediaGenerationInput?: StructuredMediaResolver;
    }
  ).resolveMediaGenerationInput;
  assert.equal(typeof resolver, 'function');

  const result = await resolver!(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      requestProperty: async () => 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123',
    },
    'video',
    {
      getCachedMediaPath: async () => '/tmp/subminer-youtube-media-cache/abc123/media.mkv',
    },
  );

  assert.equal(result?.path, '/tmp/subminer-youtube-media-cache/abc123/media.mkv');
  assert.equal(result?.source, 'youtube-cache');
  assert.equal(result?.singleResolvedStream, false);
  assert.equal(result?.inputOptions, undefined);
});

test('resolveMediaGenerationInput debug-logs sanitized YouTube cache hits', async () => {
  const resolver = (
    mediaSource as typeof mediaSource & {
      resolveMediaGenerationInput?: StructuredMediaResolver;
    }
  ).resolveMediaGenerationInput;
  assert.equal(typeof resolver, 'function');
  const logs: string[] = [];

  const result = await resolver!(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123&signature=secret',
      requestProperty: async () => 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123',
    },
    'video',
    {
      getCachedMediaPath: async () => '/tmp/subminer-youtube-media-cache/abc123/media.mkv',
      logDebug: (message) => logs.push(message),
    },
  );

  assert.equal(result?.source, 'youtube-cache');
  assert.match(logs.join('\n'), /kind=video source=youtube-cache/);
  assert.match(
    logs.join('\n'),
    /input=local:\/tmp\/subminer-youtube-media-cache\/abc123\/media\.mkv/,
  );
  assert.match(logs.join('\n'), /current=remote:www\.youtube\.com/);
  assert.doesNotMatch(logs.join('\n'), /signature=secret|videoplayback/);
});

test('resolveMediaGenerationInput does not fall back to direct remote streams when cache is required', async () => {
  const resolver = (
    mediaSource as typeof mediaSource & {
      resolveMediaGenerationInput?: StructuredMediaResolver;
    }
  ).resolveMediaGenerationInput;
  assert.equal(typeof resolver, 'function');

  const result = await resolver!(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      requestProperty: async () => 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123',
    },
    'video',
    {
      getCachedMediaPath: async () => null,
      remoteCacheMode: 'required',
    },
  );

  assert.equal(result, null);
});

test('resolveMediaGenerationInput falls back when optional cache lookup fails', async () => {
  const resolver = (
    mediaSource as typeof mediaSource & {
      resolveMediaGenerationInput?: StructuredMediaResolver;
    }
  ).resolveMediaGenerationInput;
  assert.equal(typeof resolver, 'function');

  const result = await resolver!(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      requestProperty: async () => 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123',
    },
    'video',
    {
      getCachedMediaPath: async () => {
        throw new Error('cache unavailable');
      },
    },
  );

  assert.equal(result?.path, 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123');
  assert.equal(result?.source, 'stream-open-filename');
});

test('resolveMediaGenerationInput debug-logs sanitized required-cache misses', async () => {
  const resolver = (
    mediaSource as typeof mediaSource & {
      resolveMediaGenerationInput?: StructuredMediaResolver;
    }
  ).resolveMediaGenerationInput;
  assert.equal(typeof resolver, 'function');
  const logs: string[] = [];

  const result = await resolver!(
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123&signature=secret',
      requestProperty: async () => 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123',
    },
    'video',
    {
      getCachedMediaPath: async () => null,
      remoteCacheMode: 'required',
      logDebug: (message) => logs.push(message),
    },
  );

  assert.equal(result, null);
  assert.match(logs.join('\n'), /kind=video source=cache-miss/);
  assert.match(logs.join('\n'), /mode=required/);
  assert.match(logs.join('\n'), /current=remote:www\.youtube\.com/);
  assert.doesNotMatch(logs.join('\n'), /signature=secret|videoplayback|googlevideo/);
});
