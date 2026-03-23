import assert from 'node:assert/strict';
import test from 'node:test';

import { resolveMediaGenerationInputPath } from './media-source';

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
  const edlSource = [
    'edl://!new_stream;!no_clip;!no_chapters;%70%https://audio.example/videoplayback?mime=audio%2Fwebm',
    '!new_stream;!no_clip;!no_chapters;%69%https://video.example/videoplayback?mime=video%2Fmp4',
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

  assert.equal(audioResult, 'https://audio.example/videoplayback?mime=audio%2Fwebm');
  assert.equal(videoResult, 'https://video.example/videoplayback?mime=video%2Fmp4');
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
